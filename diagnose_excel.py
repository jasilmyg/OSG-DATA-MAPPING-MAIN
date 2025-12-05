"""
Diagnostic script to identify Excel file column structure
Run this to understand what columns your RBM file actually has
"""
import pandas as pd
import sys

def check_excel_file(filename):
    print(f"\n{'='*70}")
    print(f"Analyzing: {filename}")
    print(f"{'='*70}")
    
    try:
        # Try reading with different engines
        try:
            df = pd.read_excel(filename, engine='openpyxl')
            print("✓ Loaded with openpyxl engine")
        except:
            try:
                df = pd.read_excel(filename)
                print("✓ Loaded with default engine")
            except Exception as e:
                print(f"✗ Failed to load: {e}")
                return None
        
        # Show basic info
        print(f"\n📊 DataFrame Info:")
        print(f"   Rows: {len(df)}")
        print(f"   Columns: {len(df.columns)}")
        
        # Show column names and types
        print(f"\n📋 Column Details:")
        for i, col in enumerate(df.columns, 1):
            dtype = df[col].dtype
            non_null = df[col].notna().sum()
            print(f"   {i}. '{col}' (type: {dtype}, non-null: {non_null}/{len(df)})")
        
        # Show first few rows
        print(f"\n📝 First 5 rows:")
        print(df.head(5).to_string())
        
        # Check for common column name variations
        print(f"\n🔍 Column Name Analysis:")
        possible_store_cols = [col for col in df.columns if 'store' in str(col).lower() or 'branch' in str(col).lower()]
        possible_rbm_cols = [col for col in df.columns if 'rbm' in str(col).lower() or 'manager' in str(col).lower()]
        
        if possible_store_cols:
            print(f"   Possible Store columns: {possible_store_cols}")
        else:
            print(f"   ⚠️ No columns found matching 'store' or 'branch'")
            
        if possible_rbm_cols:
            print(f"   Possible RBM columns: {possible_rbm_cols}")
        else:
            print(f"   ⚠️ No columns found matching 'rbm' or 'manager'")
        
        return df
        
    except Exception as e:
        print(f"✗ Error: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    # Check all master files
    files = [
        "RBM,BDM,BRANCH.xlsx",
        "myG All Store.xlsx",
        "Future Store List.xlsx"
    ]
    
    results = {}
    for file in files:
        df = check_excel_file(file)
        if df is not None:
            results[file] = df
    
    # Summary
    print(f"\n{'='*70}")
    print("SUMMARY & RECOMMENDATIONS")
    print(f"{'='*70}")
    
    if "RBM,BDM,BRANCH.xlsx" in results:
        rbm_df = results["RBM,BDM,BRANCH.xlsx"]
        rbm_cols = list(rbm_df.columns)
        
        print("\n🎯 For RBM,BDM,BRANCH.xlsx:")
        print(f"   Current columns: {rbm_cols}")
        
        # Check what needs to be fixed
        has_store = 'Store' in rbm_cols
        has_branch = 'Branch' in rbm_cols
        has_rbm = 'RBM' in rbm_cols
        
        if has_store and has_rbm:
            print("   ✓ File is correctly formatted!")
        else:
            print("\n   ⚠️ File needs these columns:")
            if not (has_store or has_branch):
                print("      - Add 'Store' or 'Branch' column with store names")
            if not has_rbm:
                print("      - Add 'RBM' column with RBM names")
            
            # Suggest column mapping
            print("\n   💡 Suggested fix: Rename these columns:")
            for col in rbm_cols:
                if 'store' in col.lower() or 'branch' in col.lower():
                    print(f"      '{col}' → 'Store'")
                elif 'rbm' in col.lower() or 'manager' in col.lower():
                    print(f"      '{col}' → 'RBM'")
