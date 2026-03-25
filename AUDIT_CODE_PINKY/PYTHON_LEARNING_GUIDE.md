# Python & Web Development Learning Guide
## Based on Your PharmaTrack Application

---

## 🎯 **PART 1: PYTHON FUNDAMENTALS IN YOUR CODE**

### **1. Imports (Lines 26-60)**

```python
import csv           # For reading/writing CSV files
import json          # For converting data to JSON format
import os            # For file system operations (paths, directories)
import pandas as pd  # POWERFUL library for working with data (tables)
from flask import Flask, request, render_template  # Web framework
```

**What's an Import?** → Loading a library/module that adds functionality to Python

---

## 📊 **PART 2: DATA STRUCTURES USED IN YOUR CODE**

### **A. DICTIONARIES (`dict`) - Key-Value Pairs**

```python
# Line 125-128: This is a DICTIONARY (dict)
bin_to_proc = dict(zip(bin_df['BIN'], bin_df['Processor']))

# A dict looks like:
# {"000123": "CAREMARK", "002456": "OPTUMRX", "007890": "CVS"}
#   ↑KEY       ↑VALUE

# This maps BIN numbers to Processor names
```

**Dictionary Usage in Your Code:**
- Storing key-value pairs (like a real dictionary)
- Fast lookup: `bin_to_proc["000123"]` → returns "CAREMARK"

**Other Dictionaries in your code:**
```python
# Line 115-120: Created a dictionary with multiple data types
_JOB_CACHE[job_id] = {
    "paths": {...},              # nested dict
    "pharmacy_name": "ABC Pharmacy",   # string value
    "date_range": "2024-01-01 to 2024-01-31",  # string
    "summary": summary           # another dict
}
```

---

### **B. LISTS (`list`) - Ordered Collections**

```python
# Line 112-117: A LIST of dictionaries
unmapped_bins = [
    {'bin': '000123', 'rx_count': 45},
    {'bin': '002456', 'rx_count': 12},
    {'bin': '007890', 'rx_count': 8},
]

# Lists use square brackets []
# Access items by position: unmapped_bins[0] → {'bin': '000123', 'rx_count': 45}
```

**Other Lists in your code:**
```python
# Line 181-192: Pre-selected sheet names as a list
preselected_sheets = [
    "Processed Data",
    "Needs to be ordered - All",
    "Do Not Order - ALL",
    # ... more items
]

# Line 195: Create a list by extracting from a dictionary
processors = [bp['processor'] for bp in by_processor]
# This uses a LIST COMPREHENSION (explained below)
```

---

### **C. PANDAS DataFrames - Tables (2D Data)**

```python
# Line 85-86: Load a CSV file into a DataFrame
log_df = pd.read_csv(custom_log_path, dtype=str)

# A DataFrame is like an Excel spreadsheet:
# +--------+-------+--------+--------+
# | Rx #   | Drug  | Price  | Status |
# +--------+-------+--------+--------+
# | 123456 | Amox  | 25.50  | Paid   |
# | 123457 | Lipitor| 45.00  | Trans  |
# +--------+-------+--------+--------+

# Access a column: log_df['Drug']        → Series of all drugs
# Access a row:    log_df.iloc[0]        → First row as Series
# Add new column:  log_df['NewCol'] = ... → Adds a column
```

**DataFrame Operations in Your Code:**
```python
# Line 93-96: Add new columns if missing
for c in ['Ins Paid Plan 1', 'Ins Paid Plan 2', 'Plan 1 BIN', 'Plan 2 BIN']:
    if c not in log_df.columns:
        log_df[c] = 0 if 'Paid' in c else ''

# Line 100-108: Complex DataFrame operations
log_df['Winning_BIN'] = np.where(
    log_df.get('Ins Paid Plan 1', 0) >= log_df.get('Ins Paid Plan 2', 0),
    log_df.get('Plan 1 BIN', ''),
    log_df.get('Plan 2 BIN', '')
)
# This uses np.where() → Similar to IF/ELSE but for entire columns
```

---

## 🔄 **PART 3: IMPORTANT OPERATIONS**

### **String Methods**

```python
# .strip()  → Remove whitespace from beginning/end
pharmacy_name = (request.form.get('pharmacy_name') or '').strip()

# .lower() → Convert to lowercase
kinray_file.filename.lower().endswith('.csv')

# .replace() → Replace text
bin_df['BIN'].str.replace(r'\D', '', regex=True)
# This removes all non-digits (0-9) from the BIN column
```

---

### **List/Dict Comprehensions - Powerful Python Features**

```python
# Line 54-57: List comprehension - compact way to create lists
missing = [k for k, f in {
    'custom_log':  custom_log_file,
    'kinray_file': kinray_file,
}.items() if not f or not f.filename]

# Breaks down as: [k for k, f in dict.items() if CONDITION]
#                  ↑ what to add to list
#                           ↑ where to get it from
#                                           ↑ filter condition

# Equivalent longer version:
missing = []
for k, f in {'custom_log': custom_log_file, 'kinray_file': kinray_file}.items():
    if not f or not f.filename:
        missing.append(k)
```

---

### **Groupby Operations - Summarizing Data**

```python
# Line 161-166: Group and aggregate
by_processor = [
    {"processor": r['Processor'], "rx_count": int(r['rx_count']), ...}
    for _, r in grp.iterrows()
]

# grp is the result of groupby:
grp = (log_df.dropna(subset=['Processor'])
             .groupby('Processor', as_index=False)
             .agg(rx_count=('Winning_BIN', 'count'),
                  total_paid=('Winning_Paid', 'sum')))

# This groups by Processor and sums their prescriptions/payments
```

---

## 🌐 **PART 4: WEB DEVELOPMENT WITH PYTHON & FLASK**

### **What is Flask?**
Flask is a **web framework** - it helps you:
1. **Receive HTTP requests** from users (via browser)
2. **Process the data** (Python code)
3. **Send responses back** (HTML pages, JSON data)

### **Routes (@app.route)**

```python
# Line 80-81: HOMEPAGE ROUTE
@app.route('/')
def index():
    return render_template('index.html')
    # When user visits http://localhost:5000/
    # This shows the HTML page from 'templates/index.html'

# Line 83: UPLOAD ROUTE
@app.route('/upload', methods=['POST'])
def upload_file():
    # When user POSTS files to http://localhost:5000/upload
    # This function runs and processes the files
    # methods=['POST'] = only accepts POST requests (file uploads)

# Line 325: EMAIL ROUTE
@app.route('/email', methods=['POST'])
def email_report():
    # When user sends email form data to http://localhost:5000/email
    # This function sends the email
```

**HTTP Methods:**
- `GET` → Retrieve data (browser request)
- `POST` → Send data (form submission, file upload)
- `PUT` → Update data
- `DELETE` → Remove data

---

### **Request & Form Data**

```python
# Line 84-85: Get form data from user
pharmacy_name = (request.form.get('pharmacy_name') or '').strip()
date_range = (request.form.get('date_range') or '').strip()

# request.form.get() retrieves data from HTML form
# 'pharmacy_name' = name attribute in <input name="pharmacy_name">

# Line 97-104: Get uploaded files
custom_log_file = request.files.get('custom_log')
kinray_file = request.files.get('kinray_file')
# request.files.get() retrieves uploaded files
```

---

### **JSON Responses**

```python
# Line 302-306: Return JSON (data format for APIs)
return jsonify({
    "ok": True,
    "job_id": job_id,
    "summary": summary
})

# jsonify() converts Python dict to JSON format
# Frontend JavaScript receives this and updates the page
# JSON Example:
# {
#   "ok": true,
#   "job_id": "abc123xyz",
#   "summary": {...}
# }
```

---

### **File Upload & Storage**

```python
# Line 116-117: Generate unique folder for each upload
job_id = uuid4().hex         # Create unique ID (random)
job_dir = os.path.join(updir, job_id)  # Create path
os.makedirs(job_dir, exist_ok=True)    # Create the directory

# Line 125-130: Save uploaded files
custom_log_path = os.path.join(job_dir, 'custom_log.csv')
custom_log_file.save(custom_log_path)  # Save file to disk
```

---

## 📝 **PART 5: WORKFLOW OF YOUR APPLICATION**

```
1. USER VISITS http://localhost:5000/
   ↓
2. Flask runs @app.route('/') → Shows HTML form (index.html)
   ↓
3. USER UPLOADS 3 CSV FILES + pharmacy name + date range
   ↓
4. HTML Form POSTS data to http://localhost:5000/upload
   ↓
5. Flask runs @app.route('/upload') function:
   ✓ Validates files (checks if required files present)
   ✓ Creates unique folder with uuid4()
   ✓ Saves files to disk
   ✓ Loads CSV files into DataFrames
   ✓ Maps BINs to Processors using dict
   ✓ Groups data by processor
   ✓ Returns JSON with summary
   ↓
6. Frontend shows review modal with summary
   ↓
7. USER FINISHES & CLICKS "PROCESS"
   ↓
8. Flask processes data and generates Excel report
   ↓
9. USER DOWNLOADS OR EMAILS REPORT
```

---

## 🔧 **NUMPY & PANDAS OPERATIONS**

### **NumPy - Numerical Computing**

```python
# Line 103-104: np.where() - conditional logic on arrays
log_df['Winning_BIN'] = np.where(
    log_df.get('Ins Paid Plan 1', 0) >= log_df.get('Ins Paid Plan 2', 0),
    log_df.get('Plan 1 BIN', ''),      # if TRUE, use this
    log_df.get('Plan 2 BIN', '')       # if FALSE, use this
)

# Equivalent: for each row, if Plan1 >= Plan2, use Plan1 BIN, else use Plan2 BIN
```

### **Pandas - Data Manipulation**

```python
# .dropna() - remove rows with missing values
log_df.dropna(subset=['Processor'])

# .groupby() - group rows by column value
log_df.groupby('Processor', as_index=False).agg(
    rx_count=('Winning_BIN', 'count'),
    total_paid=('Winning_Paid', 'sum')
)

# .fillna() - replace missing values
log_df['column'].fillna(0)  # Replace NaN with 0

# .astype() - change data type
log_df['BIN'].astype(str)   # Convert to string
```

---

## 🎓 **KEY PYTHON CONCEPTS YOU NEED TO KNOW**

| Concept | Example | Use |
|---------|---------|-----|
| **Variable** | `name = "John"` | Store data |
| **List** | `[1, 2, 3]` | Ordered collection |
| **Dict** | `{"name": "John"}` | Key-value pairs |
| **String Methods** | `"text".lower()` | Manipulate text |
| **List Comprehension** | `[x*2 for x in range(5)]` | Create lists concisely |
| **Function** | `def my_func():` | Reusable code blocks |
| **Lambda** | `lambda x: x*2` | Anonymous function |
| **Conditionals** | `if x > 5:` | Decision making |
| **Loops** | `for x in list:` | Repeat actions |

---

## 💡 **LEARNING PATH FOR WEB DEVELOPMENT**

1. **Python Basics** (1-2 weeks)
   - Variables, strings, lists, dicts
   - Functions, loops, conditionals
   - File I/O

2. **Advanced Python** (1-2 weeks)
   - List/dict comprehensions
   - Error handling (try/except)
   - Modules & imports

3. **Flask Web Framework** (2-3 weeks)
   - Routes and decorators
   - Request/response handling
   - Templates (HTML)
   - Static files (CSS, JS)

4. **Databases** (2 weeks)
   - SQLAlchemy ORM
   - Database design
   - SQL basics

5. **Data Processing** (2 weeks)
   - Pandas DataFrames
   - NumPy arrays
   - Data cleaning

6. **Deployment** (1 week)
   - Environment variables
   - Server configuration
   - Docker containers

---

## 🚀 **NEXT STEPS**

Review the specific sections of `app$.py` and try to:
1. Identify all the **dictionaries**, **lists**, and **DataFrames**
2. Understand what each **route** (@app.route) does
3. Trace the data flow from upload → processing → download

This will build your understanding of web development with Python!

