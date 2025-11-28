import re
import pandas as pd
from collections import defaultdict
from flask import Flask, request, render_template, send_file
from io import BytesIO

app = Flask(__name__)

# ---------------- SKU CATEGORY MAPPING ----------------
sku_category_mapping = {
    "Warranty : Water Cooler/Dispencer/Geyser/RoomCooler/Heater": [
        "COOLER", "DISPENCER", "GEYSER", "ROOM COOLER", "HEATER", "WATER HEATER", "WATER DISPENSER"
    ],
    "Warranty : Fan/Mixr/IrnBox/Kettle/OTG/Grmr/Geysr/Steamr/Inductn": [
        "FAN", "MIXER", "IRON BOX", "KETTLE", "OTG", "GROOMING KIT", "GEYSER", "STEAMER",
        "INDUCTION", "CEILING FAN", "TOWER FAN", "PEDESTAL FAN", "INDUCTION COOKER",
        "ELECTRIC KETTLE", "WALL FAN", "MIXER GRINDER", "CELLING FAN"
    ],
    "AC : EWP : Warranty : AC": ["AC", "AIR CONDITIONER", "AC INDOOR"],
    "HAEW : Warranty : Air Purifier/WaterPurifier": ["AIR PURIFIER", "WATER PURIFIER"],
    "HAEW : Warranty : Dryer/MW/DishW": ["DRYER", "MICROWAVE OVEN", "DISH WASHER", "MICROWAVE OVEN-CONV"],
    "HAEW : Warranty : Ref/WM": [
        "REFRIGERATOR", "WASHING MACHINE", "WASHING MACHINE-TL", "REFRIGERATOR-DC",
        "WASHING MACHINE-FL", "WASHING MACHINE-SA", "REF", "REFRIGERATOR-CBU",
        "REFRIGERATOR-FF", "WM"
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
        "VACUUM CLEANER", "FAN", "MASSAGER", "IRON BOX", "CEILING FAN",
        "TOWER FAN", "PEDESTAL FAN", "WALL FAN", "ROBO VACCUM CLEANER"
    ],
    "AC AMC": ["AC", "AC INDOOR"]
}

# ---------------- HELPERS ----------------

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


# ---------------- ROUTES ----------------

@app.route("/")
def upload_page():
    return render_template("upload.html")


@app.route("/process", methods=["POST"])
def process():

    osg_file = request.files["osg_file"]
    product_file = request.files["product_file"]

    osg_df = pd.read_excel(osg_file)
    product_df = pd.read_excel(product_file)

    # Ensure columns exist
    for col in ['Category','Model','Customer Mobile','Invoice Number','Item Rate','IMEI','Brand']:
        if col not in product_df.columns:
            product_df[col] = ''

    product_df['Category'] = product_df['Category'].astype(str).str.upper()
    product_df['Customer Mobile'] = product_df['Customer Mobile'].astype(str)
    osg_df['Customer Mobile'] = osg_df.get('Customer Mobile','').astype(str)
    product_df['Invoice Number'] = product_df['Invoice Number'].astype(str)

    # -------- MODEL LOGIC --------
    def get_model(row):
        mobile = row['Customer Mobile']
        retailer_sku = str(row.get('Retailer SKU',''))

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

        if mapped_keywords:
            filtered = user_products[user_products['Category'].str.lower().isin(mapped_keywords)]
            if not filtered.empty and filtered['Model'].nunique() == 1:
                return filtered['Model'].iloc[0]

        return ''

    osg_df['Model'] = osg_df.apply(get_model, axis=1)

    # Merge category & brand
    cat_brand = product_df[['Customer Mobile','Model','Category','Brand']].drop_duplicates()
    osg_df = osg_df.merge(cat_brand, on=['Customer Mobile','Model'], how="left")

    # Pools
    invoice_pool = defaultdict(list)
    item_pool = defaultdict(list)
    imei_pool = defaultdict(list)

    for _, row in product_df.iterrows():
        key = (row['Customer Mobile'], row['Model'])
        invoice_pool[key].append(row['Invoice Number'])
        item_pool[key].append(row['Item Rate'])
        imei_pool[key].append(row['IMEI'])

    invoice_counter = defaultdict(int)
    item_counter = defaultdict(int)
    imei_counter = defaultdict(int)

    def pick(pool, counter, key):
        if key in pool and counter[key] < len(pool[key]):
            value = pool[key][counter[key]]
            counter[key] += 1
            return value
        return ''

    osg_df['Product Invoice Number'] = osg_df.apply(
        lambda r: pick(invoice_pool, invoice_counter, (r['Customer Mobile'], r['Model'])), axis=1)

    osg_df['Item Rate'] = osg_df.apply(
        lambda r: pick(item_pool, item_counter, (r['Customer Mobile'], r['Model'])), axis=1)

    osg_df['IMEI'] = osg_df.apply(
        lambda r: pick(imei_pool, imei_counter, (r['Customer Mobile'], r['Model'])), axis=1)

    # Final columns
    final_cols = [
        'Customer Mobile','Date','Invoice Number','Product Invoice Number','Customer Name',
        'Store Code','Branch','Region','IMEI','Category','Brand','Quantity','Item Code',
        'Model','Plan Type','EWS QTY','Item Rate','Plan Price','Sold Price','Email',
        'Product Count','Manufacturer Warranty','Retailer SKU','OnsiteGo SKU',
        'Duration (Year)','Total Coverage','Comment','Return Flag','Return against invoice No.',
        'Primary Invoice No.'
    ]

    for col in final_cols:
        if col not in osg_df.columns:
            osg_df[col] = ''

    osg_df['Quantity'] = 1
    osg_df['EWS QTY'] = 1

    output = BytesIO()
    osg_df.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="OSG_Updated.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    app.run()
