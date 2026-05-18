import pandas as pd, json, re, os

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, 'data')
os.makedirs(DATA_DIR, exist_ok=True)

DEPT_MAP = {
    '1 _ E- CIGARRETE': 'e-cigarette',
    '2 _ CIGARETTE': 'cigarette',
    '3 _ PREGNANT TEST & CONTRACEPTION': 'pregnancy'
}

def clean_name(n):
    return re.sub(r'^\[?[A-Z0-9_/]+\]?_', '', str(n)).strip()

def fmt_wa(n):
    if not n or str(n) in ['nan', '#N/A', 'None']: return None
    n = re.sub(r'[^0-9]', '', str(n))
    if not n: return None
    return 'https://wa.me/' + (n if n.startswith('62') else '62' + n.lstrip('0'))

# --- OUTLETS ---
outlets_df = pd.read_excel(os.path.join(BASE, 'data_toko.xlsx'))
outlets_df = outlets_df[outlets_df['BRANCH NAME'] != 'HEAD OFFICE'].copy()

outlets = []
for _, r in outlets_df.iterrows():
    outlets.append({
        'code': str(r['STORE CODE']),
        'name': str(r['STORE NAME']),
        'city': str(r['KAB/KOTA']),
        'region': str(r['BRANCH NAME']).replace('REGIONAL ', ''),
        'address': str(r['ALAMAT1']) if str(r['ALAMAT1']) != 'nan' else None,
        'maps': str(r['Google Maps']) if str(r['Google Maps']) not in ['nan', '#N/A'] else None,
        'wa': fmt_wa(r['No hp'])
    })

with open(os.path.join(DATA_DIR, 'outlets.json'), 'w', encoding='utf-8') as f:
    json.dump(outlets, f, ensure_ascii=False, indent=2)
print(f'outlets.json: {len(outlets)} stores')

# --- STOCK + PRODUCTS ---
stock_df = pd.read_excel(os.path.join(BASE, 'stock_data.xlsx'), sheet_name='Gabungan')
valid = stock_df[
    stock_df['Dept'].isin(DEPT_MAP.keys()) &
    (stock_df['Regional'] != 'HEAD OFFICE')
].copy()

valid['clean_name'] = valid['Product Name'].apply(clean_name)
valid['category'] = valid['Dept'].map(DEPT_MAP)

prod_df = valid.groupby('Product Code').first().reset_index()
products = []
for _, r in prod_df.iterrows():
    products.append({
        'code': str(r['Product Code']),
        'name': r['clean_name'],
        'category': r['category'],
        'image': None
    })

with open(os.path.join(DATA_DIR, 'products.json'), 'w', encoding='utf-8') as f:
    json.dump(products, f, ensure_ascii=False, indent=2)
print(f'products.json: {len(products)} products')

stocked = valid[valid['Closing'] > 0][['Product Code', 'KD Store', 'Closing']].copy()
stock_map = {}
for _, r in stocked.iterrows():
    code = str(r['Product Code'])
    if code not in stock_map:
        stock_map[code] = []
    stock_map[code].append({'store_code': str(r['KD Store']), 'qty': int(r['Closing'])})

with open(os.path.join(DATA_DIR, 'stock.json'), 'w', encoding='utf-8') as f:
    json.dump(stock_map, f, ensure_ascii=False, indent=2)
print(f'stock.json: {len(stock_map)} products with stock')
print('Done.')
