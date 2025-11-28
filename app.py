import re
import pandas as pd
from collections import defaultdict
from flask import Flask, request, render_template, send_file
from io import BytesIO

app = Flask(__name__)

# SKU Category Mapping
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

    # Case: Dur : X+Y
    match = re.search(r'Dur\s*:\s*(\d+)\+(\d+)', sku)
    if match:
        return int(match.group(1)), int(match.group(2))

    # Case: X+Y SDP-Z
    match = re.search(r'(\d+)\+(\d+)\s*SDP-(\d+)', sku)
    if match:
        manufacturer = int(match.group(1))
        sdp = match.group(3)
        remaining = match.group(2)
        return manufacturer, f"{sdp}P+{remaining}W"

    # Case: Dur : X
    match = re.search(r'Dur\s*:\s*(\d+)', sku)
    if match:
        return 1, int(match.group(1))

    # Case: fallback X+Y
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

@app.route("/")
def upload_page():
    return render_template("upload.html")

@app.route("/process", methods=["POST"])
def process():
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

    # Model logic
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

    # Assign Model
    osg_df['Model'] = osg_df.apply(get_model, axis=1)

    # Merge Category and Brand
    category_brand_df = product_df[['Customer Mobile', 'Model', 'Category', 'Brand']].drop_duplicates()
    osg_df = osg_df.merge(category_brand_df, on=['Customer Mobile', 'Model'], how='left')

    # Pools for assignment
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

    # Assign extra fields
    osg_df['Product Invoice Number'] = osg_df.apply(lambda row: assign_from_pool(row, invoice_pool, invoice_usage_counter), axis=1)
    osg_df['Item Rate'] = osg_df.apply(lambda row: assign_from_pool(row, itemrate_pool, itemrate_usage_counter), axis=1)
    osg_df['IMEI'] = osg_df.apply(lambda row: assign_from_pool(row, imei_pool, imei_usage_counter), axis=1)

    # ✅ Extract Store Code from Product Invoice Number
    osg_df['Store Code'] = osg_df['Product Invoice Number'].astype(str).apply(
        lambda x: re.search(r'\b([A-Z]{2,})\b', x).group(1) if re.search(r'\b([A-Z]{2,})\b', x) else ''
    )

    # Apply Warranty and Duration extraction
    osg_df[['Manufacturer Warranty', 'Duration (Year)']] = osg_df['Retailer SKU'].apply(
        lambda sku: pd.Series(extract_warranty_duration(sku))
    )

    # Ensure all required columns are present and in correct order
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

    # Set constant values
    osg_df['Quantity'] = 1
    osg_df['EWS QTY'] = 1
    osg_df = osg_df[final_columns]

    # Save with highlight
    output = BytesIO()
    styled = osg_df.style.apply(highlight_row, axis=1)
    styled.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="OSG_Updated_Output_With_Category_Brand.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run()
