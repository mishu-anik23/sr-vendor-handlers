import pandas as pd
import re

# Optional: Extend this list as needed
known_brands = [
    "OKF", "ASIAN CHOICE", "AFROASE", "INDIA GATE", "TRS", "HEERA", "SCHANI", "SHAN",
    "ASHK", "ASHOKA", "MAMA", "MDH", "LAZIZA", "LIJJAT", "LAILA", "ROYAL THAI RICE", "AROY-D",
    "PRAN", "MILKIS", "INDOMIE", "NONGSHIM", "WABU", "FOCO", "V-FRESH", "PLUVERA", "SPRING HOME",
    "ROYAL THAI", "ROYAL ORIENT", "RUCHI", "BOMBAY", "OVALTINE", "KULFI ICE", "AASHIRVAAD", "AHMED",
    "DABUR", "GITS", "HALDIRAM", "KURKURE", "LAYS", "MDH", "NIDO", "PATAK", "PG TIPS", "QARSHI", "AKASH",
    "GOLESTAN", "KNORR", "KTC", "MAGGI", "REGAL", "RUBICON", "SALANTY", "SHEZAN", "TAPAL DANEDAR", "TILDA",
    "NATURINDA", "JAZZA", "MTR", "ANNAM", "WESTCOAST", "RADHUNI", "LEXUS", "DAN", "IDEAL", "BIK", "BICANO",
    "RICO", "KATO", "BIBIGO", "LITTLE MOONS", "WMD", "DOUX", "HUMZA", "WEIKFIELD", "HEMANI", "ENCONA", "JH FOODS",
    "SUNRISE", "ISPAHANI", "CROWN FARM", "SHODESH", "RN BRAND", "HEER", "PARLE", "GOLDEN MOUNTAIN", "MILO",
    "NESTLÉ", "TATA", "VITAL", "TG", "MEGACHEF", "PRB", "KIKKOMAN", "LACTASOY", "PRB", "GREEN FARM", "RAITIP",
    "JHFOODS", "LIPTON", "ANNY", "SAHIBA", "BAMBOO TREE", "FARMER", "ACECOOK", " WAI WAI THAILAND", "PILLSBURY",
    "PARACHUTE", "MYM", "RENUKA", "YEN NHUNG", "CARNATION", "WABU", "KHANUM", "OYAKATA", "PRIMA", " KAIJAE",
    "HORLICKS", "HERITAGE AFRIKA", "OISHI", "VIMTO", "ML SQUID", "KOH-KAE", "COCK", "DETTOL", "MAMA'S CHOICE",
    "MAE KRUA", "SAGIKO", "SUNFLOWER", "GREEN TABLE", "JONGGA", "VIET NAM", "SHANA", "UPASTRY", "NOODLE HOUSE",
    "TAKIS", "ELEFANT", "HYPER MALT", "HOT CHIP", "SZU SHEN PO", "AGARBATTI", "NARCISSUS", "EVERBEST", "HEALTHY BOY",
    "KINGZEST", "HAIDILAO", "MAO XIONG", "HERBEX", "BAMBOO TREE", "PRESIDENT", "JIADUOBAO", "HIKARI MISO", "LAKOVO",
    "YOPOKKI", "CYPRESSA", "MEHEK", "GINGERBON", "PULMUONE", "WOK FOODS", "BAIJIA", "WEI LIH", "GOLD KILI",
    "PCD", "PRB", "SLINMY", "KAILO", "HUNG PHAT", "SHAN WAI", "HERR'S", " PATAK'S", "POR KWAN", "CARABAO", "WEIJUTE",
    "COFE", "RABBIT", "OTAFUKU", "KHONG DO", "TAMANOI", "JIA BRAND", "SHAN WAI", "YANCO", "HUNG PHAT", "MP", "WANT WANT",
    "YAN LONG", "PAN", "JING YI GEN", "GOGI", "YUANFU BRAND", "MAI WA", "SUKINA", "RAFHAN", "KIMHO", "MOGUMOGU",
    "YUM YUM", "CHIU CHOW", "TARO", "MEGA", "HAOHAO", "JUB JUB", "SKYBIRD", "MEIJI H.PANDA", "ROYAL TIGER", "SAMYANG",
    "COCON", "GENKI RAMUNE", "LAO GAN MA", "IFAD", "LKK", "YAMASA", "JONGGA", "LONGLIFE", "TONGYI", "MAN TANG XIAN",
    "MAE NAPA", "ELEPHANT", "CHUPA CHUPS", "MARUKOME", "EAGLOBE", "SQUID", "MAO XIONG", "EFP", "HEINZ", "BINGGRAE",
    "MINI MELTS", "SICHUAN WANG", "SEMPIO", "NITTAYA", "A", "ATOOM", "HENG SHUN", "JIABAO", "AQUAPEARL", "PERFIT",
    "BRITANNIA", "CROWN", "FLYING GOOSE"
]

def is_all_caps(text):
    return text.isupper() and any(c.isalpha() for c in text)


def normalize_size_ranges(name):
    """Convert shrimp sizes like 16/20 → 16-20"""
    return re.sub(r'(\d+)\s*/\s*(\d+)', r'\1-\2', name)

def format_inhalt(inhalt):
    """
    Format Inhalt: convert to uppercase and add space between number and unit.
    Example: "500g" -> "500 G", "1kg" -> "1 KG"
    """
    if not inhalt or pd.isna(inhalt):
        return ""
    
    inhalt_clean = str(inhalt).strip().upper()
    # Add space between number and unit (e.g., "500g" -> "500 G")
    inhalt_clean = re.sub(r'(\d)([A-Za-z])', r'\1 \2', inhalt_clean)
    return inhalt_clean

def parse_product_info(text):
    if pd.isna(text):
        return {"brand": None, "product_name": None, "quantity": None, "full_name": None}

    original_text = text.strip()

    # 1. Normalize spaces and commas
    text_clean = re.sub(r'\s+', ' ', original_text).strip()
    text_clean = text_clean.rstrip(',')

    # 2. Normalize shrimp size ranges (16/20 → 16-20)
    text_clean = normalize_size_ranges(text_clean)

    # 3. Extract quantity (strictly looking for x or X patterns)
    quantity_match = re.search(r'(\d+\s*[xX*]\s*\d+\s*\w*|\d+\s*\w+\s*[xX*]\s*\d+)', text_clean)
    quantity = quantity_match.group(0).strip() if quantity_match else None

    # 4. Remove quantity temporarily
    text_wo_quantity = text_clean.replace(quantity, '').strip() if quantity else text_clean

    # 5. Detect brand (now checks anywhere in text, not just start/end)
    brand = None
    for brand_candidate in known_brands:
        pattern = r'\b' + re.escape(brand_candidate) + r'\b'
        if re.search(pattern, text_wo_quantity, flags=re.IGNORECASE):
            brand = brand_candidate
            text_wo_quantity = re.sub(pattern, '', text_wo_quantity, flags=re.IGNORECASE).strip()
            break

    # 6. Clean up product name
    product_name = text_wo_quantity.strip()
    if is_all_caps(product_name):
        product_name = product_name.title()

    # Format full_name with only product weight (no carton quantity)
    moq, weight = split_quantity(quantity)
    weight_formatted = format_inhalt(weight) if weight else ""
    
    # Build full_name = brand + product_name + weight (without carton quantity)
    if weight_formatted:
        full_name = f'{brand or ""} {product_name} {weight_formatted}'.strip()
    else:
        full_name = f'{brand or ""} {product_name}'.strip()
    
    return {
        "brand": brand,
        "product_name": product_name,
        "quantity": quantity,
        "full_name": full_name
    }

def split_quantity(quantity):
    if not quantity:
        return (None, None)
    #match = re.match(r'(\d+)\s*[xX]\s*(\d+\s*\w+)', quantity)
    #if match:
     #   return match.group(1), match.group(2).upper()
    #return None, None
    # Match common patterns: 10x1kg, 500g*12, 2 x 5 KG, etc.
    match = re.match(r'(?i)(\d+)\s*[xX*]\s*(\d+\s*\w+)|(\d+\s*\w+)\s*[xX*]\s*(\d+)', quantity)
    if match:
        if match.group(1) and match.group(2):
            return (match.group(1), match.group(2).lower())
        elif match.group(3) and match.group(4):
            return (match.group(4), match.group(3).lower())
    return (None, None)

# Load original Excel
df = pd.read_excel('./sr-products.xlsx')


# Remove 'Name' column
#if 'Name' in df.columns:
    #df.drop(columns=['Name'], inplace=True)

# Remove 'inhalt' column
#if 'inhalt' in df.columns:
    #df.drop(columns=['inhalt'], inplace=True)

#if 'Stuer' in df.columns:
    #df.drop(columns=['Steur'], inplace=True)


# Parse item information
parsed = df['item'].apply(parse_product_info).apply(pd.Series)
print(parsed)

# Split quantity into MoQ and Weight
df['moq'], df['product_weight'] = zip(*parsed['quantity'].apply(split_quantity))

# Insert 'product_name' column right after 'item'
item_index = df.columns.get_loc('item')
df.insert(item_index + 1, 'product_name', parsed['product_name'])

# Insert 'product_weight' column right after 'product_name'
weight_index = df.columns.get_loc('product_name')
df.insert(weight_index + 1, 'product_weight', df.pop('product_weight'))

full_name_index = df.columns.get_loc('product_weight')
df.insert(full_name_index+1, 'full_name', parsed['full_name'])

# Insert 'MoQ' after 'Stuer'
#if 'stuer' in df.columns:
sub_category_index = df.columns.get_loc('Sub-Category')
df.insert(sub_category_index + 1, 'moq', df.pop('moq'))

# Update 'vendor' column if empty
if 'brand' in df.columns:
    df['brand'] = df['brand'].fillna(parsed['brand'])
else:
    df['brand'] = parsed['brand']

# Remove unnamed columns (those with "Unnamed")
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

# Fill empty barcodes for products with the same item (item_name) value
# Find barcode column (check common column names)
barcode_col = None
for col in ['barcode', 'Barcode', 'BARCODE', 'EAN', 'UPC', 'Barcode/EAN']:
    if col in df.columns:
        barcode_col = col
        break

if barcode_col:
    # Group by 'item' and fill empty barcodes with the barcode from rows that have it
    def fill_barcode(group):
        # Find non-empty barcode values in this group
        # Convert to string to handle both NaN and empty strings
        group_barcodes = group[barcode_col].astype(str)
        non_empty_barcodes = group_barcodes[(group_barcodes != 'nan') & (group_barcodes.str.strip() != '')]
        
        if len(non_empty_barcodes) > 0:
            # Use the first non-empty barcode value
            barcode_value = non_empty_barcodes.iloc[0]
            # Fill empty barcodes (both NaN and empty strings) in this group
            mask = (group[barcode_col].isna()) | (group[barcode_col].astype(str).str.strip() == '') | (group[barcode_col].astype(str) == 'nan')
            group.loc[mask, barcode_col] = barcode_value
        
        return group
    
    # Apply to each group of items (item_name)
    df = df.groupby('item', group_keys=False).apply(fill_barcode).reset_index(drop=True)
    print(f"✅ Filled empty barcodes based on matching item (item_name) values")

duplicates = df[df.duplicated(subset=['Art No', 'item', 'product_name'], keep=False)]
print(duplicates)
duplicates.to_excel('duplicated_cleaned_product_list.xlsx', index=False)

# Save to new Excel
df.to_excel('cleaned_sr_products.xlsx', index=False)




















































print("✅ Cleaned Excel saved as 'cleaned_product_list.xlsx'")
print("✅ Duplicated Products saved as 'duplicated_cleaned_product_list.xlsx'")

# Create a dataframe with ONLY unique rows
unique_df = df.drop_duplicates(subset=['Art No', 'item', 'product_name'], keep='first')
print(unique_df)

# Save unique rows to a new Excel file
unique_df.to_excel('sr-unique_products.xlsx', index=False)

print("✅ Unique products saved as 'sr-unique_products.xlsx'")
