# All Flask routes: /, /upload, /finalize, /email, /review, /download, /pick_folder; holds _JOB_CACHE.

import mimetypes
import os
import re
import shutil
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from glob import glob
from urllib.parse import quote
from uuid import uuid4

import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog

from flask import (Blueprint, current_app, jsonify, make_response, redirect,
                   render_template, request, send_file, send_from_directory,
                   url_for)
from werkzeug.utils import secure_filename

from processing.log_parser import _filter_custom_log_transmitted_paid_ins
from processing.pipeline import process_custom_log_data

bp = Blueprint('main', __name__)

# keep small "job context" between /upload -> /review -> /finalize
_JOB_CACHE = {}  # { job_id: { "paths": {...}, "summary": {...}, "pharmacy_name":..., "date_range":... } }


@bp.route('/')
def index():
    return render_template('index.html')


@bp.route('/upload', methods=['POST'])
def upload_file():
    pharmacy_name = (request.form.get('pharmacy_name') or '').strip()
    date_range = (request.form.get('date_range') or '').strip()

    try:
        vendor_count = int(request.form.get('vendor_count', 0))
    except ValueError:
        vendor_count = 0

    # ---- FILE OBJECTS ----
    custom_log_file = request.files.get('custom_log')      # CSV
    all_pbm_file = request.files.get('all_pbm')         # CSV (optional)
    kinray_file = request.files.get('kinray_file')     # CSV
    bin_master_file = request.files.get('bin_master')      # CSV (required)

    # ---- Required uploads ----
    missing = [k for k, f in {
        'custom_log':  custom_log_file,
        'kinray_file': kinray_file,
        'bin_master':  bin_master_file,
    }.items() if not f or not f.filename]
    if missing:
        return (f"Missing required file(s): {', '.join(missing)}. "
                f"Got files: {list(request.files.keys())}", 400)

    # ---- Save uploads to per-job dir ----
    updir = current_app.config['UPLOAD_FOLDER']
    os.makedirs(updir, exist_ok=True)

    job_id = uuid4().hex
    job_dir = os.path.join(updir, job_id)
    os.makedirs(job_dir, exist_ok=True)

    custom_log_path = os.path.join(job_dir, 'custom_log.csv')
    kinray_path = os.path.join(job_dir, 'kinray.csv')
    bin_master_path = os.path.join(job_dir, 'bin_master.csv')
    all_pbm_path = None

    if not kinray_file.filename.lower().endswith('.csv'):
        return "KINRAY file must be a CSV.", 400

    custom_log_file.save(custom_log_path)
    kinray_file.save(kinray_path)
    bin_master_file.save(bin_master_path)

    if all_pbm_file and all_pbm_file.filename:
        if not all_pbm_file.filename.lower().endswith('.csv'):
            return "ALL PBM file must be a CSV.", 400
        all_pbm_path = os.path.join(job_dir, 'all_pbm.csv')
        all_pbm_file.save(all_pbm_path)

    # ---- Optional vendor files ----
    vendor_paths = []
    if vendor_count and vendor_count > 0:
        for i in range(1, vendor_count + 1):
            vf = request.files.get(f'vendor{i}_file')
            vname = (request.form.get(
                f'vendor{i}_name') or f'Vendor{i}').strip()
            if vf and vf.filename:
                safe = re.sub(r'[^A-Za-z0-9_.-]+', '_', vname) or f'Vendor{i}'
                dest = os.path.join(job_dir, f'{safe}.csv')
                vf.save(dest)
                vendor_paths.append(dest)

    # ---- Build summary for review (processors, total Rx, $ by processor) ----
    log_df = pd.read_csv(custom_log_path, dtype=str)
    log_df, _status_col, _kept_rows, _dropped_rows = _filter_custom_log_transmitted_paid_ins(log_df)
    bin_df = pd.read_csv(bin_master_path, dtype=str)
    # ✅ Add this right here
    for c in ['Ins Paid Plan 1', 'Ins Paid Plan 2', 'Plan 1 BIN', 'Plan 2 BIN']:
        if c not in log_df.columns:
            log_df[c] = 0 if 'Paid' in c else ''  # safe default

    # normalize for summary
    for c in ['Ins Paid Plan 1', 'Ins Paid Plan 2']:
        if c in log_df.columns:
            log_df[c] = pd.to_numeric(log_df[c], errors='coerce').fillna(0)

    # choose winning bin/paid per row
    log_df['Winning_BIN'] = np.where(log_df.get('Ins Paid Plan 1', 0) >= log_df.get('Ins Paid Plan 2', 0),
                                     log_df.get('Plan 1 BIN', ''), log_df.get('Plan 2 BIN', ''))
    log_df['Winning_BIN'] = log_df['Winning_BIN'].astype(
        str).str.replace(r'\D', '', regex=True).str.zfill(6)
    log_df['Winning_Paid'] = np.where(
        log_df['Winning_BIN'] == log_df.get('Plan 1 BIN', ''),
        log_df.get('Ins Paid Plan 1', 0),
        log_df.get('Ins Paid Plan 2', 0)
    )
    log_df['Winning_Group'] = np.where(log_df.get('Ins Paid Plan 1', 0) >= log_df.get('Ins Paid Plan 2', 0),
                                       log_df.get('Plan 1 Group #', ''), log_df.get('Plan 2 Group #', ''))
    log_df['Winning_PCN'] = np.where(log_df.get('Ins Paid Plan 1', 0) >= log_df.get('Ins Paid Plan 2', 0),
                                     log_df.get('Plan 1 PCN', ''), log_df.get('Plan 2 PCN', ''))
    # map BIN -> Processor
    bin_df['BIN'] = bin_df['BIN'].astype(str).str.replace(
        r'\D', '', regex=True).str.zfill(6)
    bin_df['Processor'] = bin_df['Processor'].astype(str).str.strip()
    bin_to_proc = dict(zip(bin_df['BIN'], bin_df['Processor']))
    log_df['Processor'] = log_df['Winning_BIN'].map(bin_to_proc)

    # ---- Unmapped BINs (Winning BINs with no Processor mapping) ----
    mask_unmapped = (
        log_df['Processor'].isna()
        & log_df['Winning_BIN'].astype(str).str.strip().ne('')
    )

    if 'Rx #' in log_df.columns:
        # count unique RX # per unmapped BIN (safer)
        unmapped_grp = (
            log_df.loc[mask_unmapped]
            .groupby('Winning_BIN', as_index=False)
            .agg(rx_count=('Rx #', lambda s: s.astype(str).str.strip()
                           .replace('', np.nan).dropna().nunique()))
        )
    else:
        # fallback: count rows
        unmapped_grp = (
            log_df.loc[mask_unmapped]
            .groupby('Winning_BIN', as_index=False)
            .size().rename(columns={'size': 'rx_count'})
        )

    unmapped_grp = unmapped_grp.sort_values('rx_count', ascending=False)

    unmapped_bins = [
        {'bin': r['Winning_BIN'], 'rx_count': int(r['rx_count'])}
        for _, r in unmapped_grp.iterrows()
    ]
    unmapped_total_bins = len(unmapped_bins)
    unmapped_total_rx = int(
        unmapped_grp['rx_count'].sum()) if not unmapped_grp.empty else 0

    # total Rx (unique)
    if 'Rx #' in log_df.columns:
        total_rx = (log_df['Rx #'].astype(str).str.strip().replace(
            '', np.nan).dropna().nunique())
    else:
        total_rx = len(log_df)

    grp = (log_df.dropna(subset=['Processor'])
                 .groupby('Processor', as_index=False)
                 .agg(rx_count=('Winning_BIN', 'count'),
                      total_paid=('Winning_Paid', 'sum')))
    grp = grp.sort_values('total_paid', ascending=False)

    # ✅ Insert here
    if grp.empty:
        by_processor = []
        processors = []
    else:
        by_processor = [
            {"processor": r['Processor'], "rx_count": int(
                r['rx_count']), "total_paid": float(r['total_paid'])}
            for _, r in grp.iterrows()
        ]
        processors = [bp['processor'] for bp in by_processor]

    # ✅ expose sheet choices for the modal
    sheets_available = [
        "Processed Data",
        "Needs to be ordered - All",
        "Do Not Order - ALL",
        "Never Ordered - Check",
        "Refills 0 - Call Doctor",
        "RX Comparison - All",
        "RX Comparison +ve",
        "MFP Drugs - RX",
        "Missed Refill - Revenue Recovery",
        "BIN to Processor",
        "Summary"
    ]
    # sensible defaults (you can change in UI)
    preselected_sheets = [
        "Processed Data",
        "Needs to be ordered - All",
        "Do Not Order - ALL",
        "Never Ordered - Check",
        "Refills 0 - Call Doctor",
        "MFP Drugs - RX",
        "Missed Refill - Revenue Recovery",
        "BIN to Processor",
        "Summary"
    ]

    summary = {
        "total_rx": int(log_df.shape[0]),
        "processors": sorted(log_df["Processor"].dropna().astype(str).str.strip().unique().tolist()),
        "by_processor": (
            log_df.groupby("Processor", dropna=True)
            .agg(rx_count=("Rx #", "count"), total_paid=("Winning_Paid", "sum"))
            .reset_index()
            .rename(columns={"Processor": "processor"})
            .to_dict(orient="records")
        ),
        "unmapped_bins": unmapped_bins,                # [{bin, rx_count}, ...]
        "unmapped_total_bins": unmapped_total_bins,    # e.g., 7
        "unmapped_total_rx": unmapped_total_rx,        # e.g., 128
        "note_unmapped": "Update the BIN Master file to map these BINs to processors."

    }

    # cache minimal job context for finalize
    _JOB_CACHE[job_id] = {
        "paths": {
            "job_dir": job_dir,
            "custom_log": custom_log_path,
            "kinray": kinray_path,
            "bin_master": bin_master_path,
            "all_pbm": all_pbm_path,
            "vendors": vendor_paths,
        },
        "pharmacy_name": pharmacy_name,
        "date_range": date_range,
        "summary": summary
    }

    # respond with job + summary (front-end will open the modal)
    return jsonify({
        "ok": True,
        "job_id": job_id,
        "summary": summary
    })


@bp.route('/email', methods=['POST'])
def email_report():
    from_email = request.form.get('from_email', '').strip()
    to_email = request.form.get('to_email', '').strip()
    message = request.form.get('message', '').strip()
    job_id = request.form.get('job_id', '').strip()

    if not from_email or not to_email:
        return jsonify({"ok": False, "error": "Sender and recipient emails are required."}), 400

    # Figure out which file to attach
    processed_dir = os.path.join(
        current_app.root_path, current_app.config.get('PROCESSED_FOLDER', 'processed'))
    attach_path = None

    # 1) Prefer the job's file if we stored it
    ctx = _JOB_CACHE.get(job_id) if job_id else None
    if ctx and 'outfile' in ctx and os.path.exists(ctx['outfile']):
        attach_path = ctx['outfile']
    else:
        # 2) Fallback to newest .xlsx in processed folder
        candidates = glob.glob(os.path.join(processed_dir, '*.xlsx'))
        if not candidates:
            return jsonify({"ok": False, "error": "No report file found to attach."}), 404
        attach_path = max(candidates, key=os.path.getmtime)

    # Build email with attachment
    msg = MIMEMultipart()
    msg['Subject'] = "PharmaTrack Report"
    msg['From'] = from_email
    msg['To'] = to_email
    msg.attach(MIMEText(message or "Here's your PharmaTrack report."))

    ctype, enc = mimetypes.guess_type(attach_path)
    if ctype is None:
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)

    with open(attach_path, 'rb') as f:
        part = MIMEBase(maintype, subtype)
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment',
                        filename=os.path.basename(attach_path))
        msg.attach(part)

    # Send (configure SMTP for your environment)
    try:
        # Example: Gmail SMTP (requires an app password)
        SMTP_USER = os.environ.get('SMTP_USER') or from_email
        # app password or your relay's credential
        SMTP_PASS = os.environ.get('SMTP_PASS')

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(from_email, [to_email], msg.as_string())
        return jsonify({"ok": True})
    except Exception as e:
        # Print full traceback to server logs to help debugging (will appear in the console)
        import traceback as _tb
        _tb.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


@bp.route('/review', methods=['GET'])
def review_job():
    job_id = request.args.get('job_id', '')
    ctx = _JOB_CACHE.get(job_id)
    if not ctx:
        return jsonify({"ok": False, "error": "Invalid job_id"}), 404
    return jsonify({"ok": True, "job_id": job_id, "summary": ctx["summary"]})


@bp.route('/download')
def download_file():
    # filename = request.args.get('filename', '')
    # if not filename:
    #     return "Missing filename", 400
    # # prevent path traversal
    # safe = os.path.basename(filename)
    # fullpath = os.path.join(app.root_path, app.config.get('PROCESSED_FOLDER', 'processed'), safe)
    # if not os.path.exists(fullpath):
    #     return "File not found", 404
    # return send_file(fullpath, as_attachment=True, download_name=safe)
    filename = request.args.get("filename")
    directory = os.path.join(current_app.root_path, current_app.config.get(
        'PROCESSED_FOLDER', 'processed'))
    response = make_response(send_from_directory(
        directory, filename, as_attachment=True))
    # 🚀 Remove "Internet zone" mark by using generic content-type
    response.headers["Content-Disposition"] = f'attachment; filename="{filename}"'
    response.headers["Content-Type"] = "application/octet-stream"
    response.headers["X-Content-Type-Options"] = "nosniff"
    return response


@bp.errorhandler(500)
def _handle_500(e):
    # Ensure JSON instead of HTML debugger
    return jsonify({"error": str(e)}), 500


@bp.route('/finalize', methods=['POST'])
def finalize_job():
    data = request.get_json(force=True, silent=True) or {}
    job_id = (data.get('job_id') or '').strip()

    # Validate required folder paths
    main_save_dir = (data.get('main_save_dir') or '').strip()
    audit_save_dir = (data.get('audit_save_dir') or '').strip()

    if not main_save_dir:
        return jsonify({"ok": False, "error": "Main report destination folder is required"}), 400
    if not audit_save_dir:
        return jsonify({"ok": False, "error": "Audit workbook destination folder is required"}), 400

    if not os.path.isdir(main_save_dir):
        return jsonify({"ok": False, "error": f"Main report folder does not exist: {main_save_dir}"}), 400
    if not os.path.isdir(audit_save_dir):
        return jsonify({"ok": False, "error": f"Audit workbook folder does not exist: {audit_save_dir}"}), 400

    selected_processors = [str(p).strip().upper() for p in (
        data.get('selected_processors') or []) if p]
    selected_sheets = [str(s).strip()
                       for s in (data.get('selected_sheets') or []) if s]

    ctx = _JOB_CACHE.get(job_id)
    if not ctx:
        return jsonify({"ok": False, "error": "Invalid job_id"}), 404

    paths = ctx["paths"]
    processed_dir = os.path.join(
        current_app.root_path, current_app.config.get('PROCESSED_FOLDER', 'processed'))
    os.makedirs(processed_dir, exist_ok=True)

    try:
        # include Kinray first, then the extras
        vendor_paths = [paths["kinray"], *paths["vendors"]]

        result = process_custom_log_data(
            custom_log_path=paths["custom_log"],
            all_pbm_path=paths["all_pbm"],
            bin_master_path=paths["bin_master"],
            vendor_paths=vendor_paths,
            pharmacy_name=ctx.get("pharmacy_name", ""),
            date_range=ctx.get("date_range", ""),
            selected_processors=selected_processors,
            selected_sheets=selected_sheets,
            user_audit_dir=audit_save_dir,
        )
        if not result or not isinstance(result, dict):
            return jsonify({"ok": False, "error": "Report generation returned no filename"}), 500

        main_name = result.get("main")
        audit_name = result.get("audit")
        #print(main_name, audit_name)
        if not main_name:
            return jsonify({"ok": False, "error": "Missing main filename"}), 500
        # IMPORTANT: do NOT change the name returned by the writer
        #fullpath = os.path.join(processed_dir, out_filename)
        processed_dir = os.path.join(
            current_app.root_path, current_app.config.get('PROCESSED_FOLDER', 'processed'))
        main_path = os.path.join(processed_dir, main_name)
        audit_path = os.path.join(processed_dir, audit_name)

        if not os.path.exists(main_path):
            # print("[finalize] processed dir listing:", os.listdir(processed_dir))
            return jsonify({"ok": False, "error": "Main output file not found"}), 500


        ctx["outfile_main"] = main_path

        # Copy main file to user-specified folder (now required)
        try:
            import shutil
            main_dest = os.path.join(main_save_dir, main_name)
            shutil.copy2(main_path, main_dest)
            #print(f"[finalize] copied main report to: {main_dest}")
        except Exception as e:
            #print(f"[finalize] error copying main file to {main_save_dir}: {e}")
            return jsonify({"ok": False, "error": f"Failed to copy main report: {str(e)}"}), 500

        audit_exists = False
        if audit_name:
            audit_path = os.path.join(processed_dir, audit_name)
            if os.path.exists(audit_path):
                ctx["outfile_audit"] = audit_path
                audit_exists = True
                # Copy audit file to user-specified folder (now required)
                try:
                    import shutil
                    audit_dest = os.path.join(audit_save_dir, audit_name)
                    shutil.copy2(audit_path, audit_dest)
                    #(f"[finalize] copied audit report to: {audit_dest}")
                except Exception as e:
                    #print(f"[finalize] error copying audit file to {audit_save_dir}: {e}")
                    return jsonify({"ok": False, "error": f"Failed to copy audit report: {str(e)}"}), 500
            else:
                print(f"[finalize] audit file expected but not found: {audit_path}")


        # if not os.path.exists(fullpath):
        #     # (Optional) one fallback try: a regex-sanitized variant
        #     alt = re.sub(r'[^A-Za-z0-9()._\-\s]+', '_', out_filename)
        #     altpath = os.path.join(processed_dir, alt)
        #     if os.path.exists(altpath):
        #         out_filename = alt
        #     else:
        #         print("[finalize] expected at:", os.path.abspath(fullpath))
        #         print("[finalize] processed dir listing:",
        #               os.listdir(processed_dir))
        #         return jsonify({"ok": False, "error": f"Output file not found: {out_filename}"}), 500

        # ctx['outfile'] = fullpath

        # return jsonify({
        #     "ok": True,
        #     "filename": out_filename,
        #     "download_url": f"/download?filename={quote(out_filename)}"
        # })
        # return jsonify({
        #     "ok": True,
        #     "filename": main_name,
        #     "download_url": f"/download?filename={quote(main_name)}",
        #     # new behaviour (audit report)
        #     # "audit_filename": audit_name,
        #     # "audit_download_url": f"/download?filename={quote(audit_name)}",
        # })
        resp = {
            "ok": True,
            "main_filename": main_name,
            "main_download_url": f"/download?filename={quote(main_name)}",
            "download_url": f"/download?filename={quote(main_name)}"  # backwards-compatible
        }
        if audit_name and ctx.get("outfile_audit"):
            resp.update({
                "audit_filename": audit_name,
                "audit_download_url": f"/download?filename={quote(audit_name)}"
            })
        return jsonify(resp)


    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@bp.route('/pick_folder', methods=['GET'])
def pick_folder():
    """Open a native folder picker on the server (local desktop) and return the chosen path.

    Note: This only works safely when the Flask app runs on the user's desktop (local machine).
    """
    try:
        # Use a transient Tk root to show the directory dialog
        root = tk.Tk()
        root.withdraw()
        path = filedialog.askdirectory(title='Choose folder to save audit workbook')
        root.destroy()
        if not path:
            return jsonify({"ok": False, "error": "No folder selected"}), 400
        return jsonify({"ok": True, "path": path})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
