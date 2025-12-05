import pandas as pd

# Check RBM file columns
print("=" * 60)
print("Checking RBM,BDM,BRANCH.xlsx")
print("=" * 60)
try:
    rbm_df = pd.read_excel("RBM,BDM,BRANCH.xlsx", engine='openpyxl')
    print(f"✓ File loaded successfully")
    print(f"  Rows: {len(rbm_df)}")
    print(f"  Columns: {list(rbm_df.columns)}")
    print(f"\nFirst 3 rows:")
    print(rbm_df.head(3))
except Exception as e:
    print(f"✗ Error loading file: {e}")

print("\n" + "=" * 60)
print("Checking myG All Store.xlsx")
print("=" * 60)
try:
    store_df = pd.read_excel("myG All Store.xlsx", engine='openpyxl')
    print(f"✓ File loaded successfully")
    print(f"  Rows: {len(store_df)}")
    print(f"  Columns: {list(store_df.columns)}")
    print(f"\nFirst 3 rows:")
    print(store_df.head(3))
except Exception as e:
    print(f"✗ Error loading file: {e}")

print("\n" + "=" * 60)
print("Checking Future Store List.xlsx")
print("=" * 60)
try:
    future_df = pd.read_excel("Future Store List.xlsx", engine='openpyxl')
    print(f"✓ File loaded successfully")
    print(f"  Rows: {len(future_df)}")
    print(f"  Columns: {list(future_df.columns)}")
    print(f"\nFirst 3 rows:")
    print(future_df.head(3))
except Exception as e:
    print(f"✗ Error loading file: {e}")
