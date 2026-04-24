PharmaTrack – Pharmacy Data Processing and Audit System
Overview

PharmaTrack is a backend-driven application built to process pharmacy data and generate structured Excel reports for audit and decision-making.

It helps pharmacy owners understand:

Which drugs are causing financial loss
Which insurance is underpaying
What inventory needs to be reordered

The system is already used in real pharmacy operations.

Problem

Pharmacies handle data from multiple sources:

Custom log files (billing data)
Vendor purchase files (Kinray, McKesson, others)
Insurance BIN mapping

This data is usually analyzed manually in Excel, which leads to:

Errors in calculations
No clear visibility into losses
Time-consuming audit preparation
Solution

PharmaTrack automates the full workflow:

Processes multiple input files
Applies pricing and insurance logic
Generates structured Excel reports with clear insights
Features
Data Processing
Accepts custom log (CSV)
Accepts multiple vendor files (Excel)
Accepts BIN mapping file (CSV)
Business Logic
Identifies winning insurance (Plan 1 vs Plan 2)
Calculates:
Quantity billed (Q)
Insurance paid (P)
Difference (D)
Computes package-level billing using drug package size
Pricing Logic
Extracts latest vendor price per NDC
Uses date-based fallback if price is missing
Calculates final cost using quantity filled
Output Reports

Generated Excel includes:

Insurance-wise quantity, paid, and difference columns
Vendor purchase comparison
Total purchased vs billed

Sheets included:

Needs to be Ordered – identifies shortage using negative difference
Do Not Order – excess inventory cases
Never Ordered – billed but not purchased drugs
Insurance from BIN Master – mapped insurance data
Architecture
User Upload
   ↓
Flask Backend (routes)
   ↓
Processing Layer (Pandas)
   ↓
Pricing and BIN Mapping Logic
   ↓
Business Rule Engine
   ↓
Excel Generation (OpenPyXL)
   ↓
Downloadable Report
Tech Stack
Python
Flask
Pandas
OpenPyXL
HTML (basic frontend)
Project Structure
routes/        → API endpoints
processing/    → core data processing logic
utils/         → helper functions
excel/         → Excel generation logic
reports/       → output files
Templates/     → frontend templates
config.py      → configuration
run.py         → application entry point
Usage
Upload the following files:
Custom Log (CSV)
Vendor Files (Excel)
BIN Mapping File (CSV)
Click submit
Download the generated Excel report
Real-World Impact
Used in multiple pharmacy stores
Reduces manual audit effort
Identifies loss-making drugs
Improves purchasing decisions
