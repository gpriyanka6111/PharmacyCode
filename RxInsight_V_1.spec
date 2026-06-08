# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

block_cipher = None

# ── Collect all data files needed ──
datas = [
    ('templates',   'templates'),    # Jinja2 HTML templates
    ('static',      'static'),       # CSS, JS, images, favicon
    ('processing',  'processing'),   # Python package — needed for importlib
    ('excel',       'excel'),        # Excel generation modules
    ('routes',      'routes'),       # Flask blueprint routes
    ('utils',       'utils'),        # Helper utilities
    ('config.py',   '.'),            # App config
    ('app$.py',     '.'),            # Main app file (dollar sign name)
]

# Include BIN master or any .csv/.xlsx data files in root if they exist
for fname in os.listdir('.'):
    if fname.endswith(('.csv', '.xlsx', '.xls')) and os.path.isfile(fname):
        datas.append((fname, '.'))

# ── Hidden imports ──
hiddenimports = [
    'app$',
    # Flask ecosystem
    'flask',
    'flask.templating',
    'flask.json',
    'jinja2',
    'jinja2.ext',
    'werkzeug',
    'werkzeug.utils',
    'werkzeug.routing',
    # FlaskWebGUI
    'flaskwebgui',
    'psutil',
    # Data processing
    'pandas',
    'pandas._libs.tslibs.base',
    'numpy',
    'numpy.core._multiarray_umath',
    # Excel
    'openpyxl',
    'openpyxl.styles',
    'openpyxl.utils',
    'openpyxl.worksheet',
    'openpyxl.formatting',
    'openpyxl.worksheet.table',
    'openpyxl.chart',
    'xlsxwriter',
    # GUI
    'tkinter',
    'tkinter.filedialog',
    # Standard lib
    'importlib',
    'importlib.util',
    'email',
    'email.mime',
    'email.mime.multipart',
    'email.mime.text',
    'email.mime.base',
    'smtplib',
    # App modules
    'routes.main',
    'processing.pipeline',
    'processing.vendor_parser',
    'processing.log_parser',
    'processing.all_pbm_parser',
    'processing.kinray_pricing',
    'excel.formatting',
    'excel.order_sheets',
    'excel.support_sheets',
    'excel.rx_comparison_sheets',
    'excel.refill_sheets',
    'excel.summary_sheet',
    'excel.audit_workbook',
    'excel.processed_data_sheet',
    'utils.helpers',
]

a = Analysis(
    ['run.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'IPython',
        'jupyter',
        'notebook',
        'pytest',
        'test',
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='RxInsight V.1',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='rxinsight.ico',           # Your new icon
)
