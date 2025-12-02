# Report 1 Error Fix - KeyError: "['Store'] not in index"

## Problem Summary
The application was throwing a KeyError at line 268 when trying to merge the RBM DataFrame:
```python
.merge(rbm_df[['Store', 'RBM']], on='Store', how='left')
```

## Root Cause
The RBM master file (`RBM,BDM,BRANCH.xlsx`) was being loaded, but the column normalization logic wasn't properly handling all scenarios:
- The code assumed the column was named 'Branch' and tried to rename it to 'Store'
- If the file already had a 'Store' column, or if it had a different column name, the rename operation failed silently
- When the merge operation tried to access `rbm_df[['Store', 'RBM']]`, the 'Store' column didn't exist, causing the KeyError

## Solution Applied
Added robust column validation and normalization:

1. **Added diagnostic logging** to show what columns exist before and after normalization
2. **Conditional rename** - Only renames 'Branch' to 'Store' if 'Branch' exists and 'Store' doesn't
3. **Validation checks** - Verifies that both 'Store' and 'RBM' columns exist after normalization
4. **Clear error messages** - If required columns are missing, returns a descriptive error with available columns

## Code Changes
**File**: `app.py`
**Lines**: 177-179 → 177-195

### Before:
```python
book1_df['DATE'] = pd.to_datetime(book1_df['DATE'], dayfirst=True, errors='coerce')
book1_df = book1_df.dropna(subset=['DATE'])
rbm_df.rename(columns={'Branch': 'Store'}, inplace=True)
```

### After:
```python
book1_df['DATE'] = pd.to_datetime(book1_df['DATE'], dayfirst=True, errors='coerce')
book1_df = book1_df.dropna(subset=['DATE'])

# Normalize RBM DataFrame columns
print(f"RBM file columns before normalization: {rbm_df.columns.tolist()}", file=sys.stderr)
if 'Branch' in rbm_df.columns and 'Store' not in rbm_df.columns:
    rbm_df.rename(columns={'Branch': 'Store'}, inplace=True)
    print("Renamed 'Branch' to 'Store' in RBM file", file=sys.stderr)

# Validate RBM file has required columns
if 'Store' not in rbm_df.columns:
    print(f"ERROR: RBM file missing 'Store' column. Available columns: {rbm_df.columns.tolist()}", file=sys.stderr)
    return f"Error: RBM file must have 'Store' or 'Branch' column. Found columns: {rbm_df.columns.tolist()}", 400
if 'RBM' not in rbm_df.columns:
    print(f"ERROR: RBM file missing 'RBM' column. Available columns: {rbm_df.columns.tolist()}", file=sys.stderr)
    return f"Error: RBM file must have 'RBM' column. Found columns: {rbm_df.columns.tolist()}", 400

print(f"RBM file columns after normalization: {rbm_df.columns.tolist()}", file=sys.stderr)
```

## Testing Recommendations
1. Check the server logs to see what columns are in your RBM file
2. Verify that the RBM file has the correct column names: either 'Store' or 'Branch', plus 'RBM'
3. If the error persists, the logs will now show you exactly what columns are available

## Next Steps
If you still encounter issues, please share:
- The column names shown in the error message
- A screenshot of the first few rows of your `RBM,BDM,BRANCH.xlsx` file
