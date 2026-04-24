# PharmaTrack – Pharmacy Data Processing and Audit System

## Overview
PharmaTrack is a data processing application built to analyze pharmacy billing and purchase data. It generates structured Excel reports that help pharmacy owners understand financial performance and inventory needs.

The system processes multiple input files and applies business rules to identify losses, compare vendor purchases, and support audit workflows.

---

## Problem
Pharmacies work with data from different sources:
- Billing logs (custom log files)
- Vendor purchase files (Kinray and others)
- Insurance BIN mapping

This data is usually handled manually in Excel, which leads to:
- Calculation errors  
- No clear visibility into losses  
- Time-consuming audit preparation  

---

## Solution
PharmaTrack automates the workflow:
- Ingests multiple data sources
- Applies pricing and insurance logic
- Generates structured Excel reports with clear outputs

---

## Features

### Input Handling
- Custom Log (CSV)
- Vendor Files (Excel)
- BIN Mapping File (CSV)

### Data Processing
- Standardizes vendor file formats
- Maps insurance using BIN values
- Identifies winning insurance (Plan 1 vs Plan 2)

### Calculations
- Quantity billed (Q)
- Insurance paid (P)
- Difference (D)
- Package-level billing using drug package size

### Pricing Logic
- Extracts latest vendor price per NDC
- Uses date-based fallback when data is missing
- Calculates cost based on quantity filled

### Output Reports
The generated Excel file includes:
- Insurance-wise quantity, paid, and difference columns
- Vendor purchase comparison
- Total purchased vs billed

Sheets generated:
- Needs to be Ordered – identifies shortages
- Do Not Order – identifies excess inventory
- Never Ordered – billed but not purchased drugs
- Insurance from BIN Master – mapped insurance data

---

## Architecture

```
User Upload
   ↓
Flask Backend (routes)
   ↓
Processing Layer (Pandas)
   ↓
Pricing and BIN Mapping Logic
   ↓
Business Rules
   ↓
Excel Generation (OpenPyXL)
   ↓
Downloadable Report
```

---

## Tech Stack
- Python
- Flask
- Pandas
- OpenPyXL
- HTML (basic frontend)

---

## Project Structure

```
routes/        API endpoints
processing/    core data processing logic
utils/         helper functions
excel/         Excel generation logic
reports/       generated output files
Templates/     frontend templates
config.py      configuration
run.py         application entry point
```

---

## Setup and Run

### Prerequisites
- Python 3.x

### Installation
```
pip install -r requirements.txt
```

### Run Application
```
python run.py
```

---

## Usage
1. Upload required files:
   - Custom Log (CSV)
   - Vendor Files (Excel)
   - BIN Mapping File (CSV)

2. Submit for processing

3. Download the generated Excel report

---

## Real-World Usage
- Used in multiple pharmacy stores
- Supports audit preparation
- Helps identify loss-making drugs
- Improves purchasing decisions
