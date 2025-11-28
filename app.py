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
    return {"status": "ok", "version": "1.5", "deployed": "2025-11-28 15:25"}

@app.route("/mapping")
def mapping_page():
    return render_template("mapping.html")

@app.route("/report1")
def report1_page():
    today = datetime.today().strftime('%Y-%m-%d')
    first_of_month = datetime.today().replace(day=1).strftime('%Y-%m-%d')
    return render_template("report1.html", today=today, first_of_month=first_of_month)

@app.route("/report2")
def report2_page():
    today = datetime.today().strftime('%Y-%m-%d')
    return render_template("report2.html", today=today)

# ---------------------------------------------------------
# PROCESS: DATA MAPPING
# ---------------------------------------------------------

@app.route("/process_mapping", methods=["POST"])
def process_mapping():
    if 'osg_file' not in request.files or 'product_file' not in request.files:
        return "Missing files", 400

    osg_file = request.files["osg_file"]
    product_file = request.files["product_file"]

    osg_df = pd.read_excel(osg_file)
    product_df = pd.read_excel(product_file)

    # Preprocess
    product_df['Category'] = product_df['Category'].astype(str).str.upper().replace('NAN', '')
    product_df['Model'] = product_df['Model'].fillna('')
    product_df['Customer Mobile'] = product_df['Customer Mobile'].astype(str)
    product_df['Invoice Number'] = product_df['Invoice Number'].astype(str)
    product_df['Item Rate'] = pd.to_numeric(product_df['Item Rate'], errors='coerce')
    product_df['IMEI'] = product_df['IMEI'].astype(str).fillna('')
    product_df['Brand'] = product_df['Brand'].fillna('')
    osg_df['Customer Mobile'] = osg_df['Customer Mobile'].astype(str)

    def get_model(row):
        mobile = row['Customer Mobile']
        retailer_sku = str(row.get('Retailer SKU', ''))
        invoice = str(row.get('Invoice Number', ''))
        user_products = product_df[product_df['Customer Mobile'] == mobile]

        if user_products.empty:
            return ''

        unique_models = user_products['Model'].dropna().unique()
        if len(unique_models) == 1:
            return unique_models[0]

        mapped_keywords = []
        for sku_key, keywords in sku_category_mapping.items():
            if sku_key in retailer_sku:
                mapped_keywords = [kw.lower() for kw in keywords]
                break

        filtered = user_products[user_products['Category'].str.lower().isin(mapped_keywords)]

        if not filtered.empty:
             if filtered['Model'].nunique() == 1:
                return filtered['Model'].iloc[0]

        slab_min, slab_max = extract_price_slab(retailer_sku)
        if slab_min and slab_max:
            slab_filtered = filtered[(filtered['Item Rate'] >= slab_min) & (filtered['Item Rate'] <= slab_max)]
            if not slab_filtered.empty and slab_filtered['Model'].nunique() == 1:
                return slab_filtered['Model'].iloc[0]

            invoice_filtered = slab_filtered[slab_filtered['Invoice Number'].astype(str) == invoice]
            if not invoice_filtered.empty and invoice_filtered['Model'].nunique() == 1:
                return invoice_filtered['Model'].iloc[0]
        return ''

    osg_df['Model'] = osg_df.apply(get_model, axis=1)

    category_brand_df = product_df[['Customer Mobile', 'Model', 'Category', 'Brand']].drop_duplicates()
    osg_df = osg_df.merge(category_brand_df, on=['Customer Mobile', 'Model'], how='left')

    invoice_pool = defaultdict(list)
    itemrate_pool = defaultdict(list)
    imei_pool = defaultdict(list)

    for _, row in product_df.iterrows():
        key = (row['Customer Mobile'], row['Model'])
        invoice_pool[key].append(row['Invoice Number'])
        itemrate_pool[key].append(row['Item Rate'])
        imei_pool[key].append(row['IMEI'])

    invoice_usage_counter = defaultdict(int)
    itemrate_usage_counter = defaultdict(int)
    imei_usage_counter = defaultdict(int)

    def assign_from_pool(row, pool, counter_dict):
        key = (row['Customer Mobile'], row['Model'])
        values = pool.get(key, [])
        index = counter_dict[key]
        if index < len(values):
            counter_dict[key] += 1
            return values[index]
        return ''

    osg_df['Product Invoice Number'] = osg_df.apply(lambda row: assign_from_pool(row, invoice_pool, invoice_usage_counter), axis=1)
    osg_df['Item Rate'] = osg_df.apply(lambda row: assign_from_pool(row, itemrate_pool, itemrate_usage_counter), axis=1)
    osg_df['IMEI'] = osg_df.apply(lambda row: assign_from_pool(row, imei_pool, imei_usage_counter), axis=1)

    osg_df['Store Code'] = osg_df['Product Invoice Number'].astype(str).apply(
        lambda x: re.search(r'\b([A-Z]{2,})\b', x).group(1) if re.search(r'\b([A-Z]{2,})\b', x) else ''
    )

    osg_df[['Manufacturer Warranty', 'Duration (Year)']] = osg_df['Retailer SKU'].apply(
        lambda sku: pd.Series(extract_warranty_duration(sku))
    )

    final_columns = [
        'Customer Mobile', 'Date', 'Invoice Number','Product Invoice Number', 'Customer Name', 'Store Code', 'Branch', 'Region',
        'IMEI', 'Category', 'Brand', 'Quantity', 'Item Code', 'Model', 'Plan Type', 'EWS QTY', 'Item Rate',
        'Plan Price', 'Sold Price', 'Email', 'Product Count', 'Manufacturer Warranty', 'Retailer SKU', 'OnsiteGo SKU',
        'Duration (Year)', 'Total Coverage', 'Comment', 'Return Flag', 'Return against invoice No.',
        'Primary Invoice No.'
    ]

    for col in final_columns:
        if col not in osg_df.columns:
            osg_df[col] = ''

    osg_df['Quantity'] = 1
    osg_df['EWS QTY'] = 1
    osg_df = osg_df[final_columns]

    output = BytesIO()
    styled = osg_df.style.apply(highlight_row, axis=1)
    styled.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="OSG_Updated_Output.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------------------------------------------------------
# PROCESS: REPORT 1 (SALES REPORT)
# ---------------------------------------------------------

@app.route("/process_report1", methods=["POST"])
def process_report1():
    try:
        print("=== START REPORT 1 PROCESSING ===")
        print("Reading form data...")
        report_date = pd.to_datetime(request.form['report_date'])
        prev_date = pd.to_datetime(request.form['prev_date'])
        print(f"Report date: {report_date}, Prev date: {prev_date}")
        
        print("Reading uploaded files...")
        curr_osg_file = request.files['curr_osg_file']
        product_file = request.files['product_file']
        prev_osg_file = request.files.get('prev_osg_file')
        print(f"Files received - OSG: {curr_osg_file.filename}, Product: {product_file.filename}, Prev: {prev_osg_file.filename if prev_osg_file else 'None'}")
        
        print("Loading master files...")
        # Load master files from backend
        future_store_df = pd.read_excel("myG All Store.xlsx")
        rbm_df = pd.read_excel("RBM,BDM,BRANCH.xlsx")
        print(f"Master files loaded - Stores: {len(future_store_df)}, RBM: {len(rbm_df)}")
        
        print("Reading OSG file...")
        
        book1_df = pd.read_excel(curr_osg_file)
        
        # Ensure required columns exist in OSG file
        if 'Branch' in book1_df.columns:
            book1_df.rename(columns={'Branch': 'Store'}, inplace=True)
        elif 'Store' not in book1_df.columns:
            return "<h1>Error</h1><p>OSG file must have either 'Branch' or 'Store' column</p>", 400
            
        if 'DATE' not in book1_df.columns and 'Date' not in book1_df.columns:
            return "<h1>Error</h1><p>OSG file must have 'DATE' or 'Date' column</p>", 400
        if 'Date' in book1_df.columns:
            book1_df.rename(columns={'Date': 'DATE'}, inplace=True)
            
        if 'QUANTITY' not in book1_df.columns and 'Quantity' not in book1_df.columns:
            book1_df['QUANTITY'] = 1
        elif 'Quantity' in book1_df.columns:
            book1_df.rename(columns={'Quantity': 'QUANTITY'}, inplace=True)
            
        if 'AMOUNT' not in book1_df.columns and 'Amount' not in book1_df.columns:
            return "<h1>Error</h1><p>OSG file must have 'AMOUNT' or 'Amount' column</p>", 400
        if 'Amount' in book1_df.columns:
            book1_df.rename(columns={'Amount': 'AMOUNT'}, inplace=True)
        
        book1_df['DATE'] = pd.to_datetime(book1_df['DATE'], dayfirst=True, errors='coerce')
        book1_df = book1_df.dropna(subset=['DATE'])
        rbm_df.rename(columns={'Branch': 'Store'}, inplace=True)

        product_df = pd.read_excel(product_file)
        
        # Ensure required columns exist in Product file
        if 'Branch' in product_df.columns:
            product_df.rename(columns={'Branch': 'Store'}, inplace=True)
        elif 'Store' not in product_df.columns:
            return "<h1>Error</h1><p>Product file must have either 'Branch' or 'Store' column</p>", 400
            
        if 'Date' in product_df.columns:
            product_df.rename(columns={'Date': 'DATE'}, inplace=True)
        elif 'DATE' not in product_df.columns:
            return "<h1>Error</h1><p>Product file must have 'DATE' or 'Date' column</p>", 400
            
        if 'Sold Price' in product_df.columns:
            product_df.rename(columns={'Sold Price': 'AMOUNT'}, inplace=True)
        elif 'Amount' in product_df.columns:
            product_df.rename(columns={'Amount': 'AMOUNT'}, inplace=True)
        elif 'AMOUNT' not in product_df.columns:
            return "<h1>Error</h1><p>Product file must have 'Sold Price', 'Amount', or 'AMOUNT' column</p>", 400
        
        product_df['DATE'] = pd.to_datetime(product_df['DATE'], dayfirst=True, errors='coerce')
        product_df = product_df.dropna(subset=['DATE'])
        if 'QUANTITY' not in product_df.columns:
            product_df['QUANTITY'] = 1

        mtd_df = book1_df[book1_df['DATE'].dt.month == report_date.month]
        today_df = mtd_df[mtd_df['DATE'].dt.date == report_date.date()]
        today_agg = today_df.groupby('Store', as_index=False).agg({'QUANTITY': 'sum', 'AMOUNT': 'sum'}).rename(columns={'QUANTITY': 'FTD Count', 'AMOUNT': 'FTD Value'})
        mtd_agg = mtd_df.groupby('Store', as_index=False).agg({'QUANTITY': 'sum', 'AMOUNT': 'sum'}).rename(columns={'QUANTITY': 'MTD Count', 'AMOUNT': 'MTD Value'})

        product_mtd_df = product_df[product_df['DATE'].dt.month == report_date.month]
        product_today_df = product_mtd_df[product_mtd_df['DATE'].dt.date == report_date.date()]
        product_today_agg = product_today_df.groupby('Store', as_index=False).agg({'QUANTITY': 'sum', 'AMOUNT': 'sum'}).rename(columns={'QUANTITY': 'Product_FTD_Count', 'AMOUNT': 'Product_FTD_Amount'})
        product_mtd_agg = product_mtd_df.groupby('Store', as_index=False).agg({'QUANTITY': 'sum', 'AMOUNT': 'sum'}).rename(columns={'QUANTITY': 'Product_MTD_Count', 'AMOUNT': 'Product_MTD_Amount'})

        if prev_osg_file and prev_osg_file.filename != '':
            prev_df = pd.read_excel(prev_osg_file)
            prev_df.rename(columns={'Branch': 'Store'}, inplace=True)
            prev_df['DATE'] = pd.to_datetime(prev_df['DATE'], dayfirst=True, errors='coerce')
            prev_df = prev_df.dropna(subset=['DATE'])
            prev_mtd_df = prev_df[prev_df['DATE'].dt.month == prev_date.month]
            prev_mtd_agg = prev_mtd_df.groupby('Store', as_index=False).agg({'AMOUNT': 'sum'}).rename(columns={'AMOUNT': 'PREV MONTH SALE'})
        else:
            prev_mtd_agg = pd.DataFrame(columns=['Store', 'PREV MONTH SALE'])

        all_stores = pd.DataFrame(pd.Series(pd.concat([future_store_df['Store'], book1_df['Store'], product_df['Store']]).unique(), name='Store'))
        report_df = all_stores.merge(today_agg, on='Store', how='left') \
                              .merge(mtd_agg, on='Store', how='left') \
                              .merge(product_today_agg, on='Store', how='left') \
                              .merge(product_mtd_agg, on='Store', how='left') \
                              .merge(prev_mtd_agg, on='Store', how='left') \
                              .merge(rbm_df[['Store', 'RBM']], on='Store', how='left')

        required_columns = ['Store', 'FTD Count', 'FTD Value', 'Product_FTD_Amount', 'MTD Count', 'MTD Value', 'Product_MTD_Amount', 'PREV MONTH SALE', 'RBM']
        for col in required_columns:
            if col not in report_df.columns:
                report_df[col] = 0
        report_df = report_df.rename(columns={'Store': 'Store Name'})

        report_df[['FTD Count', 'FTD Value', 'MTD Count', 'MTD Value', 'Product_FTD_Count', 'Product_FTD_Amount', 'Product_MTD_Count', 'Product_MTD_Amount', 'PREV MONTH SALE']] = report_df[['FTD Count', 'FTD Value', 'MTD Count', 'MTD Value', 'Product_FTD_Count', 'Product_FTD_Amount', 'Product_MTD_Count', 'Product_MTD_Amount', 'PREV MONTH SALE']].fillna(0).astype(int)

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

        excel_output = BytesIO()
        with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Formats
            colors_palette = {'primary_blue': '#1E3A8A', 'light_blue': '#DBEAFE', 'success_green': '#065F46', 'light_green': '#D1FAE5', 'warning_orange': '#EA580C', 'light_orange': '#FED7AA', 'danger_red': '#DC2626', 'light_red': '#FEE2E2', 'accent_purple': '#7C3AED', 'light_purple': '#EDE9FE', 'neutral_gray': '#6B7280', 'light_gray': '#F9FAFB', 'white': '#FFFFFF', 'dark_blue': '#0F172A', 'mint_green': '#10B981', 'light_mint': '#ECFDF5', 'royal_blue': '#3B82F6', 'light_royal': '#EBF8FF'}
            
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
                'asp_format': workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'num_format': '₹#,##0.00'}),
                'asp_format_alt': workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': colors_palette['neutral_gray'], 'bg_color': colors_palette['light_royal'], 'num_format': '₹#,##0.00'}),
                'asp_total': workbook.add_format({'bold': True, 'font_size': 12, 'font_color': colors_palette['white'], 'bg_color': colors_palette['mint_green'], 'align': 'center', 'valign': 'vcenter', 'border': 2, 'border_color': colors_palette['mint_green'], 'num_format': '₹#,##0.00'})
            }

            ist = pytz.timezone('Asia/Kolkata')
            ist_time = datetime.now(ist)

            all_data = report_df.sort_values('MTD Value', ascending=False)
            worksheet = workbook.add_worksheet("All Stores")
            headers = ['Store Name', 'FTD Count', 'FTD Value', 'FTD Value Conversion', 'MTD Count', 'MTD Value', 'MTD Value Conversion', 'PREV MONTH SALE', 'DIFF %', 'ASP']
            
            worksheet.merge_range(0, 0, 0, len(headers) - 1, "OSG All Stores Sales Report", formats['title'])
            worksheet.merge_range(1, 0, 1, len(headers) - 1, f"Report Generated: {ist_time.strftime('%d %B %Y %I:%M %p IST')}", formats['subtitle'])

            for col, header in enumerate(headers):
                worksheet.write(5, col, header, formats['header_main'])
                worksheet.set_column(col, col, 15)
            worksheet.set_column(0, 0, 30)

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
            worksheet.write(total_row, 3, "", formats['total_row'])
            worksheet.write(total_row, 4, all_data['MTD Count'].sum(), formats['total_row'])
            worksheet.write(total_row, 5, all_data['MTD Value'].sum(), formats['total_row'])
            worksheet.write(total_row, 6, "", formats['total_row'])
            worksheet.write(total_row, 7, all_data['PREV MONTH SALE'].sum(), formats['total_row'])
            worksheet.write(total_row, 8, "", formats['total_row'])
            worksheet.write(total_row, 9, "", formats['total_row'])

        excel_output.seek(0)
        return send_file(excel_output, as_attachment=True, download_name=f"OSG_Sales_Report_{datetime.now().strftime('%Y%m%d')}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print("ERROR IN REPORT 1:")
        print(error_details)
        return f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Error - OSG Sales Report</title>
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
                <h1>⚠️ Error Processing Sales Report</h1>
                <div class="error-msg">
                    <strong>Error:</strong> {str(e)}
                </div>
                <h3>Full Stack Trace:</h3>
                <pre>{error_details}</pre>
                <p><a href="/report1">← Go Back</a></p>
            </div>
        </body>
        </html>
        """, 500

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
