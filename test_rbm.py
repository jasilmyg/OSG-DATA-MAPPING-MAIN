"""
Quick test script to verify RBM file structure
This will tell you if your files are compatible with the updated code
"""

try:
    import pandas as pd
    import openpyxl
    print("[OK] Required libraries found\n")
except ImportError as e:
    print(f"[ERROR] Missing library: {e}")
    print("Install with: pip install pandas openpyxl")
    exit(1)

def test_rbm_file():
    print("="*70)
    print("TESTING: RBM,BDM,BRANCH.xlsx")
    print("="*70)
    
    try:
        # Load the file
        rbm_df = pd.read_excel("RBM,BDM,BRANCH.xlsx", engine='openpyxl')
        print(f"[OK] File loaded successfully ({len(rbm_df)} rows)\n")
        
        # Show original columns
        print(f"Original Columns: {list(rbm_df.columns)}\n")
        
        # Simulate the normalization logic
        column_mapping = {}
        store_col_found = False
        rbm_col_found = False
        
        # Check for Store/Branch column
        for col in rbm_df.columns:
            col_lower = str(col).lower().strip()
            if col_lower in ['store', 'branch', 'store name', 'branch name', 'storename', 'branchname']:
                if col != 'Store':
                    column_mapping[col] = 'Store'
                    print(f"  -> Will map '{col}' to 'Store'")
                else:
                    print(f"  -> Column 'Store' already exists")
                store_col_found = True
                break
        
        # Check for RBM column  
        for col in rbm_df.columns:
            col_lower = str(col).lower().strip()
            if col_lower in ['rbm', 'rbm name', 'rbmname', 'manager', 'regional manager', 'rm']:
                if col != 'RBM':
                    column_mapping[col] = 'RBM'
                    print(f"  -> Will map '{col}' to 'RBM'")
                else:
                    print(f"  -> Column 'RBM' already exists")
                rbm_col_found = True
                break
        
        # Apply mappings
        if column_mapping:
            rbm_df.rename(columns=column_mapping, inplace=True)
            print(f"\n[OK] Mappings applied: {column_mapping}")
        
        # Check results
        print(f"\nFinal Columns: {list(rbm_df.columns)}\n")
        
        if store_col_found and rbm_col_found:
            print("[SUCCESS] File is compatible!")
            print(f"\nSample data:")
            print(rbm_df[['Store', 'RBM']].head(5).to_string())
            return True
        else:
            print("[FAILED] Missing required columns:")
            if not store_col_found:
                print("   - Need a 'Store' or 'Branch' column")
            if not rbm_col_found:
                print("   - Need an 'RBM' or 'Manager' column")
            print(f"\nAvailable columns: {list(rbm_df.columns)}")
            print("\nRecommended action:")
            print("1. Open RBM,BDM,BRANCH.xlsx in Excel")
            print("2. Rename columns to 'Store' and 'RBM'")
            print("3. Or share the column names above so we can add support for them")
            return False
            
    except FileNotFoundError:
        print("[ERROR] File not found: RBM,BDM,BRANCH.xlsx")
        print("Make sure you're running this from the correct directory")
        return False
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_rbm_file()
    
    print("\n" + "="*70)
    if success:
        print("[SUCCESS] ALL TESTS PASSED - Ready to deploy!")
    else:
        print("[FAILED] Fix the Excel file or share column names")
    print("="*70)
