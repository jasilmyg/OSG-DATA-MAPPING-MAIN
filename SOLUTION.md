# COMPLETE SOLUTION: KeyError "['Store'] not in index"

## Problem
You're getting a KeyError when processing Report 1:
```
KeyError: "['Store'] not in index"
```

This occurs at line 283 (local) / line 380 (deployed) when trying to merge RBM data.

## Root Cause Analysis

The `RBM,BDM,BRANCH.xlsx` file doesn't have the expected column names. The code was looking for exact matches of 'Store' and 'RBM', but your file might have:
- Different casing (e.g., 'store', 'STORE', 'branch', 'BRANCH')
- Different names (e.g., 'Branch Name', 'Store Name', 'Manager', 'RBM Name')
- Extra spaces or formatting issues

## Solution Applied

I've completely rewritten the RBM file normalization logic to:

### 1. **Case-Insensitive Column Detection**
The code now searches for columns using lowercase comparison:
- For Store: Matches 'store', 'branch', 'store name', 'branch name', 'storename', 'branchname'
- For RBM: Matches 'rbm', 'rbm name', 'rbmname', 'manager', 'regional manager', 'rm'

### 2. **Comprehensive Logging**
Added detailed diagnostic output that shows:
- Original column names before normalization
- Which columns were found and how they're being mapped
- Final DataFrame shape and sample data
- Clear error messages if columns are missing

### 3. **Data Cleaning**
After normalization, the code:
- Keeps only the 'Store' and 'RBM' columns
- Strips whitespace from all values
- Removes rows with NaN values

## How to Deploy & Test

### Step 1: Deploy to Render
```bash
# Commit the changes
git add app.py
git commit -m "Fix: Enhanced RBM column normalization for Report 1"
git push origin main
```

### Step 2: Check Render Logs
After deployment, when you run Report 1, check the Render logs. You should now see:
```
==== RBM DataFrame Column Normalization ====
Original columns: ['<actual column names>']
DataFrame shape: (X, Y)
  Mapping '<original name>' -> 'Store'
  Mapping '<original name>' -> 'RBM'
Applied column mappings: {...}
...
Final RBM DataFrame: X rows with columns: ['Store', 'RBM']
```

### Step 3: If It Still Fails
If you still get an error, the logs will now show you **exactly** what columns exist in your RBM file. Share those log lines with me.

## Quick Fix for Excel File (Alternative Solution)

If you want to fix the Excel file instead of relying on code detection:

1. Open `RBM,BDM,BRANCH.xlsx`
2. Rename the columns to exactly match:
   - Column with store/branch names → **Store** (exact case)
   - Column with RBM/manager names → **RBM** (exact case)
3. Save the file
4. Re-deploy or update the file on Render

## Testing Locally (Optional)

To test this before deploying:

1. Run the diagnostic script:
```powershell
python diagnose_excel.py
```

This will show you what columns are in your Excel files.

2. If you have pandas installed, you can test the normalization:
```python
import pandas as pd

rbm_df = pd.read_excel("RBM,BDM,BRANCH.xlsx", engine='openpyxl')
print("Columns:", list(rbm_df.columns))
print(rbm_df.head())
```

## What Changed in app.py

### Before (Lines 180-194):
- Simple rename: `Branch` → `Store`
- Basic validation
- Would fail if column names didn't match exactly

### After (Lines 180-245):
- Case-insensitive column detection
- Supports multiple column name variations  
- Comprehensive logging for debugging
- Data cleaning (whitespace removal, NaN filtering)
- Clear error messages showing available columns

## Expected Behavior

### Success Case:
```
==== RBM DataFrame Column Normalization ====
Original columns: ['Branch', 'RBM']  # Or whatever your file has
DataFrame shape: (50, 2)
  Mapping 'Branch' -> 'Store'
Applied column mappings: {'Branch': 'Store'}
Columns after mapping: ['Store', 'RBM']
Final RBM DataFrame: 50 rows with columns: ['Store', 'RBM']
Sample data:
      Store          RBM
0     Store1    Manager1
1     Store2    Manager2
2     Store3    Manager1
==== End RBM Normalization ====
```

### Failure Case (with helpful error):
```
==== RBM DataFrame Column Normalization ====
Original columns: ['Location', 'Manager_Name']
DataFrame shape: (50, 2)

ERROR: RBM file missing 'Store' or 'Branch' column.
Available columns: ['Location', 'Manager_Name']
Please ensure the Excel file has a column named 'Store', 'Branch', or similar.
```

## Next Steps

1. **Deploy** the updated app.py to Render
2. **Run** Report 1 again
3. **Check** the Render logs for the diagnostic output
4. **Share** the log output with me if it still fails

The new code will either:
- ✅ Successfully detect and normalize your columns
- ❌ Show you exactly what columns exist so we can add them to the detection list

---

**File Modified**: `app.py`  
**Lines Changed**: 180-245  
**Complexity**: Enhanced  
**Testing Required**: Yes - deploy and check logs
