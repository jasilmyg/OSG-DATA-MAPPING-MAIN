import re
import pandas as pd
from collections import defaultdict
from flask import Flask, request, render_template, send_file, redirect, url_for
from io import BytesIO
from datetime import datetime
import pytz
import xlsxwriter
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import sys
import gc

app = Flask(__name__)

# ---------------------------------------------------------
# SHARED / DATA MAPPING LOGIC
# ---------------------------------------------------------

sku_category_mapping = {
    "Warranty : Water Cooler/Dispencer/Geyser/RoomCooler/Heater": [
        "COOLER", "DISPENCER", "GEYSER", "ROOM COOLER", "HEATER", "WATER HEATER", "WATER DISPENSER"
    ],
    "Warranty : Fan/Mixr/IrnBox/Kettle/OTG/Grmr/Geysr/Steamr/Inductn": [
        "FAN", "MIXER", "IRON BOX", "KETTLE", "OTG", "GROOMING KIT", "GEYSER", "STEAMER", "INDUCTION",
        "CEILING FAN", "TOWER FAN", "PEDESTAL FAN", "INDUCTION COOKER", "ELECTRIC KETTLE", "WALL FAN", "MIXER GRINDER", "CELLING FAN"
    ],
    "AC : EWP : Warranty : AC": ["AC", "AIR CONDITIONER", "AC INDOOR"],
    "HAEW : Warranty : Air Purifier/WaterPurifier": ["AIR PURIFIER", "WATER PURIFIER"],
    "HAEW : Warranty : Dryer/MW/DishW": ["DRYER", "MICROWAVE OVEN", "DISH WASHER", "MICROWAVE OVEN-CONV"],
    "HAEW : Warranty : Ref/WM": [
        "REFRIGERATOR", "WASHING MACHINE", "WASHING MACHINE-TL", "REFRIGERATOR-DC",
        "WASHING MACHINE-FL", "WASHING MACHINE-SA", "REF", "REFRIGERATOR-CBU", "REFRIGERATOR-FF", "WM"
    ],
    "HAEW : Warranty : TV": ["TV", "TV 28 %", "TV 18 %"],
    "TV : TTC : Warranty and Protection : TV": ["TV", "TV 28 %", "TV 18 %"],
    "TV : Spill and Drop Protection": ["TV", "TV 28 %", "TV 18 %"],
    "HAEW : Warranty :Chop/Blend/Toast/Air Fryer/Food Processr/JMG/Induction": [
        "CHOPPER", "BLENDER", "TOASTER", "AIR FRYER", "FOOD PROCESSOR", "JUICER", "INDUCTION COOKER"
    ],
    "HAEW : Warranty : HOB and Chimney": ["HOB", "CHIMNEY"],
    "HAEW : Warranty : HT/SoundBar/AudioSystems/PortableSpkr": [
        "HOME THEATRE", "AUDIO SYSTEM", "SPEAKER", "SOUND BAR", "PARTY SPEAKER"
    ],
    "HAEW : Warranty : Vacuum Cleaner/Fans/Groom&HairCare/Massager/Iron": [
        "VACUUM CLEANER", "FAN", "MASSAGER", "IRON BOX", "CEILING FAN", "TOWER FAN", "PEDESTAL FAN", "WALL FAN", "ROBO VACCUM CLEANER"
    ],
    "AC AMC": ["AC", "AC INDOOR"]
}

def extract_price_slab(text):
    match = re.search(r"Slab\s*:\s*(\d+)K-(\d+)K", str(text))
    if match:
        return int(match.group(1)) * 1000, int(match.group(2)) * 1000
    return None, None

def extract_warranty_duration(sku):
    sku = str(sku)
    match = re.search(r'Dur\s*:\s*(\d+)\+(\d+)', sku)
    if match:
        return int(match.group(1)), int(match.group(2))
    match = re.search(r'(\d+)\+(\d+)\s*SDP-(\d+)', sku)
    if match:
        return int(match.group(1)), f"{match.group(3)}P+{match.group(2)}W"
    match = re.search(r'Dur\s*:\s*(\d+)', sku)
    if match:
        return 1, int(match.group(1))
    match = re.search(r'(\d+)\+(\d+)', sku)
    if match:
        return int(match.group(1)), int(match.group(2))
    return '', ''

def highlight_row(row):
    missing_fields = pd.isna(row.get('Model')) or str(row.get('Model')).strip() == ''
    missing_fields |= pd.isna(row.get('IMEI')) or str(row.get('IMEI')).strip() == ''
    try:
        if float(row.get('Plan Price', 0)) < 0:
            missing_fields |= True
    except:
        missing_fields |= True
    return ['background-color: lightblue'] * len(row) if missing_fields else [''] * len(row)

# ---------------------------------------------------------
# ROUTES
# ---------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/health")
def health():
    return {"status": "ok", "version": "3.1", "deployed": "2025-12-02 14:10"}

@app.route("/mapping")
def mapping_page():
    return render_template("mapping.html")

@app.route("/report1")
def report1_page():
    return render_template("report1.html")

@app.route("/report2")
def report2_page():
    return render_template("report2.html")

# ---------------------------------------------------------
# PROCESS: REPORT 1 (SALES REPORT) - REFACTORED FOR FILES WITH RBM
# ---------------------------------------------------------

@app.route("/process_report1", methods=["POST"])
def process_report1():
    try:
        print("=== START REPORT 1 PROCESSING (v3.1 - Files with RBM columns) ===", file=sys.stderr)
        
        # Check for xlsxwriter
        try:
            import xlsxwriter
            print("xlsxwriter module found.", file=sys.stderr)
        except ImportError:
            print("CRITICAL: xlsxwriter module NOT found!", file=sys.stderr)
            return "ERROR: xlsxwriter module is missing on the server.", 500

        print("Reading form data...", file=sys.stderr)
        report_date = pd.to_datetime(request.form['report_date'])
        prev_date = pd.to_datetime(request.form['prev_date'])
        
        print("Reading uploaded files...", file=sys.stderr)
        curr_osg_file = request.files['curr_osg_file']
        product_file = request.files['product_file']
        prev_osg_file = request.files.get('prev_osg_file')
        
        # Log File Sizes
        curr_osg_file.seek(0, 2)
        osg_size = curr_osg_file.tell()
        curr_osg_file.seek(0)
        
        product_file.seek(0, 2)
        prod_size = product_file.tell()
        product_file.seek(0)
        
        print(f"File Sizes - OSG: {osg_size/1024/1024:.2f} MB, Product: {prod_size/1024/1024:.2f} MB", file=sys.stderr)
        
        # =========================================================================
        # READ OSG FILE (Expected columns: DATE, QUANTITY, BILLED (QTY), AMOUNT, RBM, Branch)
        # =========================================================================
        print("\n==== Reading OSG File ====", file=sys.stderr)
        try:
            book1_df = pd.read_excel(curr_osg_file, engine='openpyxl')
            print(f"OSG file loaded. Shape: {book1_df.shape}", file=sys.stderr)
            print(f"Original columns: {book1_df.columns.tolist()}", file=sys.stderr)
        except Exception as e:
            print(f"Error reading OSG file: {e}", file=sys.stderr)
            return f"Error reading OSG file: {e}", 500
        
        # Normalize OSG column names (flexible to handle variations)
        column_renames = {}
        for col in book1_df.columns:
            col_str = str(col).strip()
            col_lower = col_str.lower()
            
            # Date column
            if col_lower in ['date']:
                column_renames[col] = 'DATE'
            # Quantity column
            elif col_lower in ['quantity', 'qty'] and 'billed' not in col_lower:
                column_renames[col] = 'QUANTITY'
            # Billed column
            elif 'billed' in col_lower:
                column_renames[col] = 'BILLED_QTY'
            # Amount column
            elif col_lower in ['amount']:
                column_renames[col] = 'AMOUNT'
            # RBM column
            elif col_lower in ['rbm', 'manager']:
                column_renames[col] = 'RBM'
            # Branch/Store column
            elif col_lower in ['branch', 'store']:
                column_renames[col] = 'Store'
        
        if column_renames:
            book1_df.rename(columns=column_renames, inplace=True)
            print(f"Column mappings applied: {column_renames}", file=sys.stderr)
        
        print(f"Columns after normalization: {book1_df.columns.tolist()}", file=sys.stderr)
        
        # Validate required columns
        if 'Store' not in book1_df.columns:
            return f"Error: OSG file missing 'Store' or 'Branch' column. Found columns: {book1_df.columns.tolist()}", 400
        if 'DATE' not in book1_df.columns:
            return f"Error: OSG file missing 'DATE' column. Found columns: {book1_df.columns.tolist()}", 400
        if 'AMOUNT' not in book1_df.columns:
            return f"Error: OSG file missing 'AMOUNT' column. Found columns: {book1_df.columns.tolist()}", 400
        
        # Handle QUANTITY - use BILLED_QTY if QUANTITY not found
        if 'QUANTITY' not in book1_df.columns:
            if 'BILLED_QTY' in book1_df.columns:
                book1_df['QUANTITY'] = book1_df['BILLED_QTY']
                print("Using BILLED_QTY as QUANTITY", file=sys.stderr)
            else:
                book1_df['QUANTITY'] = 1
                print("No quantity column found, defaulting to 1", file=sys.stderr)
        
        # Parse dates and clean data
        book1_df['DATE'] = pd.to_datetime(book1_df['DATE'], dayfirst=True, errors='coerce')
        book1_df = book1_df.dropna(subset=['DATE'])
        book1_df['Store'] = book1_df['Store'].astype(str).str.strip()
        book1_df = book1_df[book1_df['Store'] != 'nan']
        
        # Handle RBM column from uploaded file
        has_rbm = 'RBM' in book1_df.columns
        print(f"RBM column in OSG file: {has_rbm}", file=sys.stderr)
        
        if has_rbm:
            book1_df['RBM'] = book1_df['RBM'].astype(str).str.strip()
            book1_df = book1_df[book1_df['RBM'] != 'nan']
            rbm_stores = book1_df[['Store', 'RBM']].drop_duplicates(subset=['Store']).copy()
            print(f"Extracted RBM mapping from OSG: {len(rbm_stores)} store-RBM pairs", file=sys.stderr)
            print(f"Unique RBMs: {book1_df['RBM'].unique()[:10].tolist()}", file=sys.stderr)
        
        print(f"OSG data cleaned. Final shape: {book1_df.shape}", file=sys.stderr)
        print(f"Date range: {book1_df['DATE'].min()} to {book1_df['DATE'].max()}", file=sys.stderr)
        print(f"Unique stores: {book1_df['Store'].nunique()}", file=sys.stderr)
        
        # =========================================================================
        # CALCULATE OSG AGGREGATIONS
        # =========================================================================
        print("\n==== Calculating OSG Aggregations ====", file=sys.stderr)
        
        mtd_df = book1_df[book1_df['DATE'].dt.month == report_date.month]
        today_df = mtd_df[mtd_df['DATE'].dt.date == report_date.date()]
        
        print(f"MTD records: {len(mtd_df)}, Today records: {len(today_df)}", file=sys.stderr)
        
        today_agg = today_df.groupby('Store', as_index=False).agg({
            'QUANTITY': 'sum',
            'AMOUNT': 'sum'
        }).rename(columns={'QUANTITY': 'FTD Count', 'AMOUNT': 'FTD Value'})
        
        mtd_agg = mtd_df.groupby('Store', as_index=False).agg({
            'QUANTITY': 'sum',
            'AMOUNT': 'sum'
        }).rename(columns={'QUANTITY': 'MTD Count', 'AMOUNT': 'MTD Value'})
        
        print(f"Today aggregation: {len(today_agg)} stores", file=sys.stderr)
        print(f"MTD aggregation: {len(mtd_agg)} stores", file=sys.stderr)
        
        # MEMORY CLEANUP 1
        print("Cleaning up OSG memory...", file=sys.stderr)
        all_osg_stores = set(book1_df['Store'].unique())
        del book1_df, mtd_df, today_df
        gc.collect()

        # =========================================================================
        # READ PRODUCT FILE
        # =========================================================================
        print("\n==== Reading Product File ====", file=sys.stderr)
        try:
            product_df = pd.read_excel(product_file, engine='openpyxl')
            print(f"Product file loaded. Shape: {product_df.shape}", file=sys.stderr)
            print(f"Original columns: {product_df.columns.tolist()}", file=sys.stderr)
        except Exception as e:
            print(f"Error reading Product file: {e}", file=sys.stderr)
            return f"Error reading Product file: {e}", 500
        
        # Normalize Product column names
        product_renames = {}
        for col in product_df.columns:
            col_lower = str(col).lower().strip()
            if col_lower in ['date']:
                product_renames[col] = 'DATE'
            elif col_lower in ['quantity', 'qty']:
                product_renames[col] = 'QUANTITY'
            elif col_lower in ['amount', 'sold price', 'price']:
                product_renames[col] = 'AMOUNT'
            elif col_lower in ['branch', 'store']:
                product_renames[col] = 'Store'
        
        if product_renames:
            product_df.rename(columns=product_renames, inplace=True)
            print(f"Product column mappings: {product_renames}", file=sys.stderr)
        
        # Validate Product file
        if 'Store' not in product_df.columns:
            return f"Error: Product file missing 'Store' column. Found: {product_df.columns.tolist()}", 400
        if 'DATE' not in product_df.columns:
            return f"Error: Product file missing 'DATE' column. Found: {product_df.columns.tolist()}", 400
        if 'AMOUNT' not in product_df.columns:
            return f"Error: Product file missing 'AMOUNT' column. Found: {product_df.columns.tolist()}", 400
        
        if 'QUANTITY' not in product_df.columns:
            product_df['QUANTITY'] = 1
        
        # Parse dates and clean
        product_df['DATE'] = pd.to_datetime(product_df['DATE'], dayfirst=True, errors='coerce')
        product_df = product_df.dropna(subset=['DATE'])
        product_df['Store'] = product_df['Store'].astype(str).str.strip()
        product_df = product_df[product_df['Store'] != 'nan']
        
        # =========================================================================
        # CALCULATE PRODUCT AGGREGATIONS
        # =========================================================================
        print("Calculating Product aggregations...", file=sys.stderr)
        
        product_mtd_df = product_df[product_df['DATE'].dt.month == report_date.month]
        product_today_df = product_mtd_df[product_mtd_df['DATE'].dt.date == report_date.date()]
        
        product_today_agg = product_today_df.groupby('Store', as_index=False).agg({
            'QUANTITY': 'sum',
            'AMOUNT': 'sum'
        }).rename(columns={'QUANTITY': 'Product_FTD_Count', 'AMOUNT': 'Product_FTD_Amount'})
        
        product_mtd_agg = product_mtd_df.groupby('Store', as_index=False).agg({
            'QUANTITY': 'sum',
            'AMOUNT': 'sum'
        }).rename(columns={'QUANTITY': 'Product_MTD_Count', 'AMOUNT': 'Product_MTD_Amount'})
        
        print(f"Product MTD: {len(product_mtd_agg)} stores, Today: {len(product_today_agg)} stores", file=sys.stderr)
        
        # MEMORY CLEANUP 2
        print("Cleaning up Product memory...", file=sys.stderr)
        product_stores = set(product_df['Store'].unique())
        del product_df, product_mtd_df, product_today_df
        gc.collect()

        # =========================================================================
        # READ PREVIOUS MONTH FILE (OPTIONAL)
        # =========================================================================
        prev_mtd_agg = pd.DataFrame(columns=['Store', 'PREV MONTH SALE'])
        if prev_osg_file and prev_osg_file.filename != '':
            print("\n==== Reading Previous Month OSG File ====", file=sys.stderr)
            try:
                prev_df = pd.read_excel(prev_osg_file, engine='openpyxl')
                
                # Normalize columns
                prev_renames = {}
                for col in prev_df.columns:
                    col_lower = str(col).lower().strip()
                    if col_lower in ['date']:
                        prev_renames[col] = 'DATE'
                    elif col_lower in ['amount']:
                        prev_renames[col] = 'AMOUNT'
                    elif col_lower in ['branch', 'store']:
                        prev_renames[col] = 'Store'
                
                if prev_renames:
                    prev_df.rename(columns=prev_renames, inplace=True)
                
                prev_df['DATE'] = pd.to_datetime(prev_df['DATE'], dayfirst=True, errors='coerce')
                prev_df = prev_df.dropna(subset=['DATE'])
                prev_df['Store'] = prev_df['Store'].astype(str).str.strip()
                
                prev_mtd_df = prev_df[prev_df['DATE'].dt.month == prev_date.month]
                prev_mtd_agg = prev_mtd_df.groupby('Store', as_index=False).agg({
                    'AMOUNT': 'sum'
                }).rename(columns={'AMOUNT': 'PREV MONTH SALE'})
                
                print(f"Previous month: {len(prev_mtd_agg)} stores", file=sys.stderr)
                
                # MEMORY CLEANUP 3
                del prev_df, prev_mtd_df
                gc.collect()
            except Exception as e:
                print(f"Warning: Could not process previous month file: {e}", file=sys.stderr)

        # =========================================================================
        # MERGE ALL DATA
        # =========================================================================
        print("\n==== Merging All Data ====", file=sys.stderr)
        
        # Get all unique stores
        all_stores_set = all_osg_stores.union(product_stores)
        
        # Try to add stores from master file if it exists
        try:
            future_store_df = pd.read_excel("myG All Store.xlsx", engine='openpyxl')
            store_col = None
            for col in future_store_df.columns:
                if str(col).lower().strip() in ['store', 'branch']:
                    store_col = col
                    break
            if store_col:
                master_stores = set(future_store_df[store_col].astype(str).str.strip())
                all_stores_set = all_stores_set.union(master_stores)
                print(f"Added {len(master_stores)} stores from master file", file=sys.stderr)
        except Exception as e:
            print(f"Note: Could not load store master file: {e}", file=sys.stderr)
        
        all_stores = pd.DataFrame({'Store': list(all_stores_set)})
        print(f"Total unique stores: {len(all_stores)}", file=sys.stderr)
        
        # Perform sequential merges
        report_df = all_stores \
            .merge(today_agg, on='Store', how='left') \
            .merge(mtd_agg, on='Store', how='left') \
            .merge(product_today_agg, on='Store', how='left') \
            .merge(product_mtd_agg, on='Store', how='left') \
            .merge(prev_mtd_agg, on='Store', how='left')
        
        # Merge RBM data if available from uploaded file
        if has_rbm:
            report_df = report_df.merge(rbm_stores, on='Store', how='left')
            print(f"Merged RBM data from uploaded file", file=sys.stderr)
        else:
            # Try to load from master file as fallback
            try:
                rbm_master = pd.read_excel("RBM,BDM,BRANCH.xlsx", engine='openpyxl')
                # Normalize RBM master columns
                rbm_renames = {}
                for col in rbm_master.columns:
                    col_lower = str(col).lower().strip()
                    if col_lower in ['branch', 'store']:
                        rbm_renames[col] = 'Store'
                    elif col_lower in ['rbm', 'manager']:
                        rbm_renames[col] = 'RBM'
                if rbm_renames:
                    rbm_master.rename(columns=rbm_renames, inplace=True)
                if 'Store' in rbm_master.columns and 'RBM' in rbm_master.columns:
                    rbm_master = rbm_master[['Store', 'RBM']].copy()
                    rbm_master['Store'] = rbm_master['Store'].astype(str).str.strip()
                    rbm_master['RBM'] = rbm_master['RBM'].astype(str).str.strip()
                    report_df = report_df.merge(rbm_master, on='Store', how='left')
                    print(f"Merged RBM data from master file", file=sys.stderr)
            except Exception as e:
                print(f"No RBM data available: {e}", file=sys.stderr)
                report_df['RBM'] = 'Unknown'
        
        print(f"Final merged data: {report_df.shape}", file=sys.stderr)
        
        # Ensure all required columns exist
        required_columns = ['Store', 'FTD Count', 'FTD Value', 'Product_FTD_Amount', 'MTD Count', 'MTD Value', 'Product_MTD_Amount', 'PREV MONTH SALE', 'RBM']
        for col in required_columns:
            if col not in report_df.columns:
                report_df[col] = 0
        
        # Fill NaN values and convert to integers
        report_df[['FTD Count', 'FTD Value', 'MTD Count', 'MTD Value', 'Product_FTD_Count', 'Product_FTD_Amount', 'Product_MTD_Count', 'Product_MTD_Amount', 'PREV MONTH SALE']] = \
            report_df[['FTD Count', 'FTD Value', 'MTD Count', 'MTD Value', 'Product_FTD_Count', 'Product_FTD_Amount', 'Product_MTD_Count', 'Product_MTD_Amount', 'PREV MONTH SALE']].fillna(0).astype(int)
        
        # Rename Store to Store Name for display
        report_df = report_df.rename(columns={'Store': 'Store Name'})
        
        # =========================================================================
        # CALCULATE METRICS
        # =========================================================================
        print("Calculating metrics...", file=sys.stderr)
        
        report_df['DIFF %'] = report_df.apply(
            lambda x: round(((x['MTD Value'] - x['PREV MONTH SALE']) / x['PREV MONTH SALE']) * 100, 2) if x['PREV MONTH SALE'] != 0 else 0,
            axis=1
        )
        report_df['ASP'] = report_df.apply(
            lambda x: round(x['MTD Value'] / x['MTD Count'], 2) if x['MTD Count'] != 0 else 0,
            axis=1
        )
        report_df['FTD Value Conversion'] = report_df.apply(
            lambda x: round((x['FTD Value'] / x['Product_FTD_Amount']) * 100, 2) if x['Product_FTD_Amount'] != 0 else 0,
            axis=1
        )
        report_df['MTD Value Conversion'] = report_df.apply(
            lambda x: round((x['MTD Value'] / x['Product_MTD_Amount']) * 100, 2) if x['Product_MTD_Amount'] != 0 else 0,
            axis=1
        )
        
        print("Generating Excel...", file=sys.stderr)
        excel_output = BytesIO()
        with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Styles
            colors_palette = {
                'primary_blue': '#1E3A8A', 'light_blue': '#DBEAFE', 'success_green': '#065F46', 'light_green': '#D1FAE5',
                'warning_orange': '#EA580C', 'light_orange': '#FED7AA', 'danger_red': '#DC2626', 'light_red': '#FEE2E2',
                'accent_purple': '#7C3AED', 'light_purple': '#EDE9FE', 'neutral_gray': '#6B7280', 'light_gray': '#F9FAFB',
                'white': '#FFFFFF', 'dark_blue': '#0F172A', 'mint_green': '#10B981', 'light_mint': '#ECFDF5',
                'royal_blue': '#3B82F6', 'light_royal': '#EBF8FF'
            }

            formats = {
                'title': workbook.add_format({'bold': True, 'font_size': 16, 'font_color': colors_palette['primary_blue'], 'align': 'center', 'valign': 'vcenter', 'bg_color': colors_palette['white'], 'border': 1, 'border_color': colors_palette['primary_blue']}),
                'subtitle': workbook.add_format({'bold': True, 'font_size': 12, 'font_color': colors_palette['neutral_gray'], 'align': 'center', 'valign': 'vcenter', 'italic': True}),
                'header_main': workbook.add_format({'bold': True, 'font_size': 11, 'font_color': colors_palette['white'], 'bg_color': colors_palette['primary_blue'], 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['primary_blue'], 'text_wrap': True}),
                'header_secondary': workbook.add_format({'bold': True, 'font_size': 10, 'font_color': colors_palette['primary_blue'], 'bg_color': colors_palette['light_blue'], 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['primary_blue']}),
                'data_normal': workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'bg_color': colors_palette['white']}),
                'data_alternate': workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'bg_color': colors_palette['light_gray']}),
                'data_store_name': workbook.add_format({'font_size': 10, 'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'bg_color': colors_palette['white'], 'indent': 1}),
                'data_store_name_alt': workbook.add_format({'font_size': 10, 'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'bg_color': colors_palette['light_gray'], 'indent': 1}),
                'conversion_low': workbook.add_format({'font_size': 10, 'font_color': colors_palette['danger_red'], 'bg_color': colors_palette['light_red'], 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['danger_red'], 'num_format': '0.00%', 'bold': True}),
                'conversion_green': workbook.add_format({'bold': True, 'font_size': 10, 'font_color': colors_palette['success_green'], 'bg_color': colors_palette['light_green'], 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['success_green'], 'num_format': '0.00%'}),
                'conversion_format': workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'num_format': '0.00%'}),
                'conversion_format_alt': workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'bg_color': colors_palette['light_royal'], 'num_format': '0.00%'}),
                'total_row': workbook.add_format({'bold': True, 'font_size': 11, 'font_color': colors_palette['white'], 'bg_color': colors_palette['mint_green'], 'align': 'center', 'valign': 'vcenter', 'border': 2, 'border_color': colors_palette['mint_green']}),
                'total_label': workbook.add_format({'bold': True, 'font_size': 11, 'font_color': colors_palette['white'], 'bg_color': colors_palette['mint_green'], 'align': 'center', 'valign': 'vcenter', 'border': 2, 'border_color': colors_palette['mint_green']}),
                'asp_format': workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1,  'border_color': colors_palette['neutral_gray'], 'num_format': '₹#,##0.00'}),
                'asp_format_alt': workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'bg_color': colors_palette['light_royal'], 'num_format': '₹#,##0.00'}),
                'asp_total': workbook.add_format({'bold': True, 'font_size': 12, 'font_color': colors_palette['white'], 'bg_color': colors_palette['mint_green'], 'align': 'center', 'valign': 'vcenter', 'border': 2, 'border_color': colors_palette['mint_green'], 'num_format': '₹#,##0.00'}),
                # RBM formats
                'rbm_title': workbook.add_format({'bold': True, 'font_size': 18, 'font_color': colors_palette['white'], 'bg_color': colors_palette['dark_blue'], 'align': 'center', 'valign': 'vcenter', 'border': 2, 'border_color': colors_palette['dark_blue']}),
                'rbm_subtitle': workbook.add_format({'bold': True, 'font_size': 11, 'font_color': colors_palette['dark_blue'], 'bg_color': colors_palette['light_royal'], 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['royal_blue'], 'italic': True}),
                'rbm_header': workbook.add_format({'bold': True, 'font_size': 11, 'font_color': colors_palette['white'], 'bg_color': colors_palette['royal_blue'], 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['royal_blue'], 'text_wrap': True}),
                'rbm_data_normal': workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'bg_color': colors_palette['white']}),
                'rbm_data_alternate': workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'bg_color': colors_palette['light_royal']}),
                'rbm_store_name': workbook.add_format({'font_size': 10, 'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'bg_color': colors_palette['white'], 'indent': 1}),
                'rbm_store_name_alt': workbook.add_format({'font_size': 10, 'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'bg_color': colors_palette['light_royal'], 'indent': 1}),
                'rbm_conversion_low': workbook.add_format({'font_size': 10, 'font_color': colors_palette['danger_red'], 'bg_color': colors_palette['light_red'], 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['danger_red'], 'num_format': '0.00%', 'bold': True}),
                'rbm_conversion_green': workbook.add_format({'bold': True, 'font_size': 10, 'font_color': colors_palette['success_green'], 'bg_color': colors_palette['light_green'], 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['success_green'], 'num_format': '0.00%'}),
                'rbm_conversion_format': workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'num_format': '0.00%'}),
                'rbm_conversion_format_alt': workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'bg_color': colors_palette['light_royal'], 'num_format': '0.00%'}),
                'rbm_total': workbook.add_format({'bold': True, 'font_size': 12, 'font_color': colors_palette['white'], 'bg_color': colors_palette['mint_green'], 'align': 'center', 'valign': 'vcenter', 'border': 2, 'border_color': colors_palette['mint_green']}),
                'rbm_total_label': workbook.add_format({'bold': True, 'font_size': 12, 'font_color': colors_palette['white'], 'bg_color': colors_palette['mint_green'], 'align': 'center', 'valign': 'vcenter', 'border': 2, 'border_color': colors_palette['mint_green']}),
                'rbm_summary': workbook.add_format({'bold': True, 'font_size': 10, 'font_color': colors_palette['royal_blue'], 'bg_color': colors_palette['light_royal'], 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['royal_blue']}),
                'rbm_performance': workbook.add_format({'bold': True, 'font_size': 10, 'font_color': colors_palette['white'], 'bg_color': colors_palette['accent_purple'], 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['accent_purple']}),
            }

            ist = pytz.timezone('Asia/Kolkata')
            ist_time = datetime.now(ist)

            # ALL STORES SHEET
            all_data = report_df.sort_values('MTD Value', ascending=False)
            worksheet = workbook.add_worksheet("All Stores")
            headers = ['Store Name', 'FTD Count', 'FTD Value', 'FTD Value Conversion', 'MTD Count', 'MTD Value', 'MTD Value Conversion', 'PREV MONTH SALE', 'DIFF %', 'ASP']
            
            worksheet.merge_range(0, 0, 0, len(headers) - 1, "OSG All Stores Sales Report", formats['title'])
            worksheet.merge_range(1, 0, 1, len(headers) - 1, f"Report Generated: {ist_time.strftime('%d %B %Y %I:%M %p IST')}", formats['subtitle'])

            total_stores = len(all_data)
            active_stores = len(all_data[all_data['FTD Count'] > 0])
            inactive_stores = total_stores - active_stores
            worksheet.merge_range(3, 0, 3, 1, "SUMMARY", formats['header_secondary'])
            worksheet.merge_range(3, 2, 3, len(headers) - 1, f"Total: {total_stores} | Active: {active_stores} | Inactive: {inactive_stores}", formats['data_normal'])

            # Dynamically adjust column widths
            column_widths = {}
            for i in range(len(headers)):
                try:
                    if i == 0:
                        max_len = max(all_data[headers[i]].astype(str).map(len).max(), len(headers[i])) + 2
                    else:
                        max_len = max(all_data[headers[i]].map(lambda x: len(str(x))).max() if headers[i] in all_data.columns else 0, len(headers[i])) + 2
                    column_widths[i] = max(max_len, 10)
                except KeyError:
                    column_widths[i] = len(headers[i]) + 2
                worksheet.set_column(i, i, column_widths[i])

            for col, header in enumerate(headers):
                worksheet.write(5, col, header, formats['header_main'])

            for row_idx, (_, row) in enumerate(all_data.iterrows(), start=6):
                is_alt = (row_idx - 6) % 2 == 1
                worksheet.write(row_idx, 0, row['Store Name'], formats['data_store_name_alt'] if is_alt else formats['data_store_name'])
                worksheet.write(row_idx, 1, row['FTD Count'], formats['data_alternate'] if is_alt else formats['data_normal'])
                worksheet.write(row_idx, 2, row['FTD Value'], formats['data_alternate'] if is_alt else formats['data_normal'])
                
                ftd_conv = row['FTD Value Conversion']
                fmt = formats['conversion_format_alt'] if is_alt else formats['conversion_format']
                if ftd_conv > 2: fmt = formats['conversion_green']
                elif ftd_conv < 2: fmt = formats['conversion_low']
                worksheet.write(row_idx, 3, ftd_conv/100, fmt)

                worksheet.write(row_idx, 4, row['MTD Count'], formats['data_alternate'] if is_alt else formats['data_normal'])
                worksheet.write(row_idx, 5, row['MTD Value'], formats['data_alternate'] if is_alt else formats['data_normal'])

                mtd_conv = row['MTD Value Conversion']
                fmt = formats['conversion_format_alt'] if is_alt else formats['conversion_format']
                if mtd_conv > 2: fmt = formats['conversion_green']
                elif mtd_conv < 2: fmt = formats['conversion_low']
                worksheet.write(row_idx, 6, mtd_conv/100, fmt)

                worksheet.write(row_idx, 7, row['PREV MONTH SALE'], formats['data_alternate'] if is_alt else formats['data_normal'])
                worksheet.write(row_idx, 8, f"{row['DIFF %']}%", formats['data_alternate'] if is_alt else formats['data_normal'])
                worksheet.write(row_idx, 9, row['ASP'], formats['asp_format_alt'] if is_alt else formats['asp_format'])

            # Total Row
            total_row = len(all_data) + 7
            worksheet.write(total_row, 0, 'TOTAL', formats['total_label'])
            worksheet.write(total_row, 1, all_data['FTD Count'].sum(), formats['total_row'])
            worksheet.write(total_row, 2, all_data['FTD Value'].sum(), formats['total_row'])
            total_ftd_conv = round((all_data['FTD Value'].sum() / all_data['Product_FTD_Amount'].sum()) * 100, 2) if all_data['Product_FTD_Amount'].sum() != 0 else 0
            worksheet.write(total_row, 3, f"{total_ftd_conv}%", formats['total_row'])
            worksheet.write(total_row, 4, all_data['MTD Count'].sum(), formats['total_row'])
            worksheet.write(total_row, 5, all_data['MTD Value'].sum(), formats['total_row'])
            total_mtd_conv = round((all_data['MTD Value'].sum() / all_data['Product_MTD_Amount'].sum()) * 100, 2) if all_data['Product_MTD_Amount'].sum() != 0 else 0
            worksheet.write(total_row, 6, f"{total_mtd_conv}%", formats['total_row'])
            worksheet.write(total_row, 7, all_data['PREV MONTH SALE'].sum(), formats['total_row'])
            total_diff = round(((all_data['MTD Value'].sum() - all_data['PREV MONTH SALE'].sum()) / all_data['PREV MONTH SALE'].sum()) * 100, 2) if all_data['PREV MONTH SALE'].sum() != 0 else 0
            worksheet.write(total_row, 8, f"{total_diff}%", formats['total_row'])
            total_asp = round(all_data['MTD Value'].sum() / all_data['MTD Count'].sum(), 2) if all_data['MTD Count'].sum() != 0 else 0
            worksheet.write(total_row, 9, total_asp, formats['asp_total'])

            if len(all_data) > 0:
                top_performer = all_data.iloc[0]
                insights_row = total_row + 2
                worksheet.merge_range(insights_row, 0, insights_row, len(headers) - 1, f"Top Performer: {top_performer['Store Name']} (Rs.{int(top_performer['MTD Value']):,})", formats['data_normal'])

            # RBM SHEETS (if RBM data exists)
            if 'RBM' in report_df.columns:
                rbm_headers = ['Store Name', 'MTD Value Conversion', 'FTD Value Conversion', 'MTD Count', 'FTD Count', 'MTD Value', 'FTD Value', 'PREV MONTH SALE', 'DIFF %', 'ASP']
                for rbm in report_df['RBM'].dropna().unique():
                    if str(rbm) == 'Unknown' or str(rbm) == 'nan':
                        continue
                    
                    rbm_data = report_df[report_df['RBM'] == rbm].sort_values('MTD Value', ascending=False)
                    worksheet_name = str(rbm)[:31]
                    rbm_ws = workbook.add_worksheet(worksheet_name)

                    rbm_ws.merge_range(0, 0, 0, len(rbm_headers) - 1, f"{rbm} - Sales Performance Report", formats['rbm_title'])
                    rbm_ws.merge_range(1, 0, 1, len(rbm_headers) - 1, f"Report Period: {ist_time.strftime('%B %Y')} | Generated: {ist_time.strftime('%d %B %Y %I:%M %p IST')}", formats['rbm_subtitle'])

                    rbm_total_stores = len(rbm_data)
                    rbm_active_stores = len(rbm_data[rbm_data['FTD Count'] > 0])
                    rbm_inactive_stores = rbm_total_stores - rbm_active_stores
                    rbm_total_amount = rbm_data['MTD Value'].sum()

                    rbm_ws.merge_range(3, 0, 3, 1, "PERFORMANCE OVERVIEW", formats['rbm_summary'])
                    rbm_ws.merge_range(3, 2, 3, len(rbm_headers) - 1, f"Total Stores: {rbm_total_stores} | Active: {rbm_active_stores} | Inactive: {rbm_inactive_stores} | Total Revenue: Rs.{rbm_total_amount:,}", formats['rbm_summary'])

                    if len(rbm_data) > 0:
                        best_performer = rbm_data.iloc[0]
                        rbm_ws.merge_range(4, 0, 4, len(rbm_headers) - 1, f"Best Performer: {best_performer['Store Name']} - Rs.{int(best_performer['MTD Value']):,}", formats['rbm_performance'])

                    # Dynamically adjust column widths for RBM sheets
                    rbm_column_widths = {}
                    for i in range(len(rbm_headers)):
                        try:
                            if i == 0:
                                max_len = max(rbm_data[rbm_headers[i]].astype(str).map(len).max(), len(rbm_headers[i])) + 2
                            else:
                                max_len = max(rbm_data[rbm_headers[i]].map(lambda x: len(str(x))).max() if rbm_headers[i] in rbm_data.columns else 0, len(rbm_headers[i])) + 2
                            rbm_column_widths[i] = max(max_len, 10)
                        except KeyError:
                            rbm_column_widths[i] = len(rbm_headers[i]) + 2
                        rbm_ws.set_column(i, i, rbm_column_widths[i])

                    for col, header in enumerate(rbm_headers):
                        rbm_ws.write(6, col, header, formats['rbm_header'])

                    for row_idx, (_, row) in enumerate(rbm_data.iterrows(), start=7):
                        is_alt = (row_idx - 7) % 2 == 1
                        rbm_ws.write(row_idx, 0, row['Store Name'], formats['rbm_store_name_alt'] if is_alt else formats['rbm_store_name'])
                        
                        mtd_conv = row['MTD Value Conversion']
                        fmt = formats['rbm_conversion_format_alt'] if is_alt else formats['rbm_conversion_format']
                        if mtd_conv > 2: rbm_ws.write(row_idx, 1, mtd_conv/100, formats['rbm_conversion_green'])
                        elif mtd_conv < 2: rbm_ws.write(row_idx, 1, mtd_conv/100, formats['rbm_conversion_low'])
                        else: rbm_ws.write(row_idx, 1, mtd_conv/100, fmt)

                        ftd_conv = row['FTD Value Conversion']
                        fmt = formats['rbm_conversion_format_alt'] if is_alt else formats['rbm_conversion_format']
                        if ftd_conv > 2: rbm_ws.write(row_idx, 2, ftd_conv/100, formats['rbm_conversion_green'])
                        elif ftd_conv < 2: rbm_ws.write(row_idx, 2, ftd_conv/100, formats['rbm_conversion_low'])
                        else: rbm_ws.write(row_idx, 2, ftd_conv/100, fmt)

                        data_fmt = formats['rbm_data_alternate'] if is_alt else formats['rbm_data_normal']
                        rbm_ws.write(row_idx, 3, int(row['MTD Count']), data_fmt)
                        rbm_ws.write(row_idx, 4, int(row['FTD Count']), data_fmt)
                        rbm_ws.write(row_idx, 5, int(row['MTD Value']), data_fmt)
                        rbm_ws.write(row_idx, 6, int(row['FTD Value']), data_fmt)
                        rbm_ws.write(row_idx, 7, int(row['PREV MONTH SALE']), data_fmt)
                        rbm_ws.write(row_idx, 8, f"{row['DIFF %']}%", data_fmt)
                        rbm_ws.write(row_idx, 9, row['ASP'], formats['asp_format_alt'] if is_alt else formats['asp_format'])

                    total_row = len(rbm_data) + 8
                    rbm_ws.write(total_row, 0, 'TOTAL', formats['rbm_total_label'])
                    rbm_total_mtd = round((rbm_data['MTD Value'].sum() / rbm_data['Product_MTD_Amount'].sum()) * 100, 2) if rbm_data['Product_MTD_Amount'].sum() != 0 else 0
                    rbm_ws.write(total_row, 1, f"{rbm_total_mtd}%", formats['rbm_total'])
                    rbm_total_ftd = round((rbm_data['FTD Value'].sum() / rbm_data['Product_FTD_Amount'].sum()) * 100, 2) if rbm_data['Product_FTD_Amount'].sum() != 0 else 0
                    rbm_ws.write(total_row, 2, f"{rbm_total_ftd}%", formats['rbm_total'])
                    rbm_ws.write(total_row, 3, rbm_data['MTD Count'].sum(), formats['rbm_total'])
                    rbm_ws.write(total_row, 4, rbm_data['FTD Count'].sum(), formats['rbm_total'])
                    rbm_ws.write(total_row, 5, rbm_data['MTD Value'].sum(), formats['rbm_total'])
                    rbm_ws.write(total_row, 6, rbm_data['FTD Value'].sum(), formats['rbm_total'])
                    rbm_ws.write(total_row, 7, rbm_data['PREV MONTH SALE'].sum(), formats['rbm_total'])
                    
                    total_prev = rbm_data['PREV MONTH SALE'].sum()
                    total_curr = rbm_data['MTD Value'].sum()
                    growth = round(((total_curr - total_prev) / total_prev) * 100, 2) if total_prev != 0 else 0
                    rbm_ws.write(total_row, 8, f"{growth}%", formats['rbm_total'])
                    
                    overall_asp = round(rbm_data['MTD Value'].sum() / rbm_data['MTD Count'].sum(), 2) if rbm_data['MTD Count'].sum() != 0 else 0
                    rbm_ws.write(total_row, 9, overall_asp, formats['asp_total'])

                    # Insights Section
                    insights_row = total_row + 2
                    if growth > 15:
                        rbm_ws.merge_range(insights_row, 0, insights_row, len(rbm_headers) - 1, f"Excellent Growth: {growth}% increase from previous month", formats['rbm_summary'])
                    elif growth < 0:
                        rbm_ws.merge_range(insights_row, 0, insights_row, len(rbm_headers) - 1, f"Needs Attention: {abs(growth)}% decrease from previous month", formats['rbm_summary'])
                    else:
                        rbm_ws.merge_range(insights_row, 0, insights_row, len(rbm_headers) - 1, f"Stable Performance: Less change from previous month", formats['rbm_summary'])

                    insights_row += 1
                    top_3_stores = rbm_data.head(3)
                    if len(top_3_stores) > 0:
                        top_stores_text = " | ".join([f"{store['Store Name']}: Rs.{int(store['MTD Value']):,}" for _, store in top_3_stores.iterrows()])
                        rbm_ws.merge_range(insights_row, 0, insights_row, len(rbm_headers) - 1, f"Top 3 Performers: {top_stores_text}", formats['rbm_summary'])

        excel_output.seek(0)
        return send_file(excel_output, as_attachment=True, download_name=f"OSG_Sales_Report_{datetime.now().strftime('%Y%m%d')}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print("=" * 80, file=sys.stderr)
        print("ERROR IN REPORT 1:", file=sys.stderr)
        print(error_details, file=sys.stderr)
        print("=" * 80, file=sys.stderr)
        return f"ERROR PROCESSING REPORT:\n{error_details}", 500

# ---------------------------------------------------------
# PROCESS: REPORT 2 (DAY VIEW)
# ---------------------------------------------------------

@app.route("/process_report2", methods=["POST"])
def process_report2():
    try:
        report_date = pd.to_datetime(request.form['report_date'])
        time_slot = request.form['time_slot']
        sales_file = request.files['sales_file']

        formatted_date = report_date.strftime("%d-%m-%Y")
        report_title = f"{formatted_date} EW Sale Till {time_slot}"

        future_df = pd.read_excel("Future Store List.xlsx")
        book2_df = pd.read_excel(sales_file)
        book2_df.rename(columns={'Branch': 'Store'}, inplace=True)

        agg = book2_df.groupby('Store', as_index=False).agg({
            'QUANTITY': 'sum',
            'AMOUNT': 'sum'
        })

        all_stores = pd.DataFrame(pd.concat([future_df['Store'], agg['Store']]).unique(), columns=['Store'])
        merged = all_stores.merge(agg, on='Store', how='left')
        merged['QUANTITY'] = merged['QUANTITY'].fillna(0).astype(int)
        merged['AMOUNT'] = merged['AMOUNT'].fillna(0).astype(int)

        merged = merged.sort_values(by='AMOUNT', ascending=False).reset_index(drop=True)

        total = pd.DataFrame([{
            'Store': 'TOTAL',
            'QUANTITY': merged['QUANTITY'].sum(),
            'AMOUNT': merged['AMOUNT'].sum()
        }])

        final_df = pd.concat([merged, total], ignore_index=True)
        final_df.rename(columns={'Store': 'Branch'}, inplace=True)

        wb = Workbook()
        ws = wb.active
        ws.title = "Store Report"

        ws.merge_cells('A1:C1')
        title_cell = ws['A1']
        title_cell.value = report_title
        title_cell.font = Font(bold=True, size=11, color="FFFFFF")
        title_cell.alignment = Alignment(horizontal='center')
        title_cell.fill = PatternFill("solid", fgColor="4F81BD")

        header_fill = PatternFill("solid", fgColor="4F81BD")
        data_fill = PatternFill("solid", fgColor="DCE6F1")
        red_fill = PatternFill("solid", fgColor="F4CCCC")
        total_fill = PatternFill("solid", fgColor="10B981")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        header_font = Font(bold=True, color="FFFFFF")
        bold_font = Font(bold=True, color="FFFFFF")

        for r_idx, row in enumerate(dataframe_to_rows(final_df, index=False, header=True), start=2):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                if r_idx == 2:
                    cell.fill = header_fill
                    cell.font = header_font
                elif final_df.loc[r_idx - 3, 'Branch'] == 'TOTAL':
                    cell.fill = total_fill
                    cell.font = bold_font
                elif final_df.loc[r_idx - 3, 'AMOUNT'] <= 0:
                    cell.fill = red_fill
                else:
                    cell.fill = data_fill
                cell.border = border
                cell.alignment = Alignment(horizontal='center')

        for col_idx, column_cells in enumerate(ws.columns, start=1):
            max_length = 0
            for cell in column_cells:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 2

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(output, as_attachment=True, download_name=f"Store_Summary_{formatted_date}_{time_slot}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print("ERROR IN REPORT 2:")
        print(error_details)
        return f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Error - Day View Report</title>
            <style>
                body {{ font-family: Arial, sans-serif; padding: 2rem; background: #f8f9fa; }}
                .error-container {{ background: white; padding: 2rem; border-radius: 8px; max-width: 1000px; margin: 0 auto; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
                h1 {{ color: #dc3545; }}
                pre {{ background: #f8f9fa; padding: 1rem; border-radius: 4px; overflow-x: auto; border-left: 4px solid #dc3545; }}
                .error-msg {{ color: #721c24; background: #f8d7da; padding: 1rem; border-radius: 4px; margin: 1rem 0; }}
            </style>
        </head>
        <body>
            <div class="error-container">
                <h1>⚠️ Error Processing Day View Report</h1>
                <div class="error-msg">
                    <strong>Error:</strong> {str(e)}
                </div>
                <h3>Full Stack Trace:</h3>
                <pre>{error_details}</pre>
                <p><a href="/report2">← Go Back</a></p>
            </div>
        </body>
        </html>
        """, 500

if __name__ == "__main__":
    app.run()
