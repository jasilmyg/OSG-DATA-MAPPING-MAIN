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
