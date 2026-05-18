#!/usr/bin/env python3
"""
convert.py — C Store Catalog Data Converter
Run: python3 convert.py
Input:  stock_data.xlsx (di root repo)
Output: data/products.json, data/stock.json, data/outlets.json
"""
import pandas as pd
import json
import re
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
STOCK_FILE = os.path.join(ROOT, 'stock_data.xlsx')
OUTLET_FILE = os.path.join(ROOT, 'outlets_data.xlsx')  # optional
OUT_DIR = os.path.join(ROOT, 'data')
os.makedirs(OUT_DIR, exist_ok=True)

DEPT_MAP = {
    '1 _ E- CIGARRETE': 'e-cigarette',
    '2 _ CIGARETTE': 'cigarette',
    '3 _ PREGNANT TEST & CONTRACEPTION': 'pregnancy',
}

def clean_name(name):
    """Hapus prefix sebelum _ (contoh: ONEPRICE_FORE SALT... -> FORE SALT...)"""
    return re.sub(r'^[A-Z0-9/\[\] ]+_', '', str(name)).strip()

def map_city(regional, store_code):
    r = str(regional).upper()
    c = str(store_code)
    if 'BANDUNG' in r: return 'Bandung'
    if 'SEMARANG' in r: return 'Semarang'
    if 'SURABAYA' in r: return 'Surabaya'
    if c.startswith('A'): return 'Jakarta'
    if c in ['B4002', 'B5003']: return 'Tangerang Selatan'
    if c.startswith('B'): return 'Tangerang'
    if c in ['C5003', 'C5004']: return 'Bogor'
    if c == 'C5001': return 'Bekasi'
    if c == 'C4002': return 'Karawang'
    return regional.title()

def fmt_wa(num):
    n = re.sub(r'[^0-9]', '', str(num))
    if not n or len(n) < 8: return None
    if not n.startswith('62'): n = '62' + n.lstrip('0')
    return f'https://wa.me/{n}'

# ── 1. Load stock ──────────────────────────────────────────────────────────
print(f"Loading {STOCK_FILE}...")
if not os.path.exists(STOCK_FILE):
    print(f"ERROR: {STOCK_FILE} tidak ditemukan!")
    sys.exit(1)

df = pd.read_excel(STOCK_FILE, sheet_name='Gabungan')
print(f"  Rows: {len(df):,}")

# ── 2. Build products list ─────────────────────────────────────────────────
products_raw = df[df['Dept'].isin(DEPT_MAP.keys())].drop_duplicates('Product Code').copy()
products_list = []
for _, row in products_raw.iterrows():
    products_list.append({
        'code': str(row['Product Code']),
        'name': clean_name(row['Product Name']),
        'category': DEPT_MAP[row['Dept']],
    })
print(f"  Products: {len(products_list):,}")

# ── 3. Build stock index (product_code -> city -> [store_names]) ──────────
relevant = df[
    df['Dept'].isin(DEPT_MAP.keys()) &
    (df['Regional'] != 'HEAD OFFICE') &
    (df['Closing'] > 0)
].copy()

stock_index = {}
for _, row in relevant.iterrows():
    code = str(row['Product Code'])
    store = str(row['Store'])
    city = map_city(str(row['Regional']), str(row['KD Store']))
    stock_index.setdefault(code, {}).setdefault(city, [])
    if store not in stock_index[code][city]:
        stock_index[code][city].append(store)
print(f"  Products with stock: {len(stock_index):,}")

# ── 4. Build outlets (from xlsx if available, else use hardcoded) ──────────
HARDCODED_OUTLETS = [
    ("C4008","JURANG PASTEUR","Bandung","Jl. Jurang No.51B, Pasteur, Kec. Sukajadi, Kota Bandung","https://maps.app.goo.gl/j3es9mHbhJxcZP636","6285966715459"),
    ("C5011","PLUTO","Bandung","Jl. Pluto Raya No.35 E, Margasari, Kec. Buahbatu Bandung","https://maps.app.goo.gl/rchmnSvAKfnzHeG49","6281944219564"),
    ("C5013","HARMONY","Bandung","Ruko Harmony Park No.4 Jl Ciwanstra No.86 Margahayu Endah","https://maps.app.goo.gl/mS2iE8GtMrvGyRe28","6287722724824"),
    ("C5015","KOPO PERMAI 3","Bandung","Ruko Kopo Permai, Jl. Kopo Permai III No.F8/7, Kab. Bandung","https://maps.app.goo.gl/Sc3JcjJHtwN5JoBw9","6287722724827"),
    ("C5014","TKI 3","Bandung","Jl. Taman Kopo Indah 3 No.F-12A, Kab. Bandung","https://maps.app.goo.gl/HTc1hu3s2adAU7mX7","6281949873164"),
    ("C5016","ISTANA SUDIRMAN","Bandung","Ruko Istana Sudirman Residen, Jl. Raya Cijerah No.11A","https://maps.app.goo.gl/WiLPY4Eq6GpoW6Vk8","6287700293365"),
    ("C4003","IBC","Bandung","Plaza IBCC, Jl. A. Yani Blok D3 No.11, Kota Bandung","https://maps.app.goo.gl/agPXDAcScuzK2GrW9","6285966715453"),
    ("C4006","VETERAN","Bandung","Jl. Veteran No.40, Kb. Pisang, Kec. Sumur Bandung","https://maps.app.goo.gl/aKefZX2Sm6Zg5HcJ8","6287722724796"),
    ("C5010","PELAJAR PEJUANG 45","Bandung","Jl. Pelajar Pejuang 45 No.102, Turangga, Kec. Lengkong","https://maps.app.goo.gl/71STHS7kXfCkR2Fn9","6281944219563"),
    ("C5012","CIJAGRA","Bandung","JL Cijagra 1 No.63, Cijagra Lengkong, Kota Bandung","https://maps.app.goo.gl/ZFxnAVf65zxFj2Vh9","6287722724813"),
    ("C4005","ASTANA","Bandung","Jl. Astana Anyar No.68, Cibadak, Kec. Astanaanyar, Bandung","https://maps.app.goo.gl/FFBrJtvAnQ4K6vKq6","6287722724795"),
    ("C4007","DUYUT","Bandung","Jl. Cibaduyut No.70, Kec. Bojongloa Kidul, Bandung","https://maps.app.goo.gl/KhcfbUjg9PyMHzzX6","6287722724798"),
    ("C5017","CIJAMBE","Bandung","Ruko Komplek Pasir Jati Cijambe No.49A, Kec. Ujung Berung","https://maps.app.goo.gl/3weQkiLxf5S6YB9HA","6287700293366"),
    ("D4004","ANGGREK","Semarang","Jl. Anggrek Raya No.4, Pekunden, Kec. Semarang Tengah","https://maps.app.goo.gl/v3AtV4K9rc1AFKg88","6287722724834"),
    ("D4006","BONJOL","Semarang","Jl. Imam Bonjol No.180H, Pendrikan Kidul, Semarang Tengah","https://maps.app.goo.gl/kbVVto9JGmv8CV777","6287722724838"),
    ("D4003","KEDUNGMUNDU","Semarang","Jl. Fatmawati No.3, Kedungmundu, Kec. Tembalang, Semarang","https://maps.app.goo.gl/HeZmXdQbsgraGjSH6","6287722724833"),
    ("D4007","KELUD","Semarang","Jl. Kelud Raya 15B, Petompon, Kec. Gajahmungkur, Semarang","https://maps.app.goo.gl/TbR96fWmoXigr6H48","6287854020222"),
    ("D4002","LAMPER","Semarang","Jl. Lamper Tengah No.168 A-B, Pandean Lamper, Semarang","https://maps.app.goo.gl/UwHvpWuwEfaj1pub9","6287722724832"),
    ("D4005","PLAMONGAN","Semarang","Jl. Brigjen Sudiarto No.158, Plamongan Sari, Pedurungan","https://maps.app.goo.gl/hPGAxWw9BxX4deNu6","6287722724837"),
    ("D4001","TLOGOSARI","Semarang","Jl. Soekarno Hatta No.101, Tlogosari Kulon, Pedurungan","https://maps.app.goo.gl/tCGDYnHufw93yS3M8","6287722724829"),
    ("F5004","KETINTANG","Surabaya","Ruko OP, Jl Ketintang Baru No 50 E, Kel. Ketintang, Surabaya","https://maps.app.goo.gl/Z7egzeXwwyz8w3jXA","628175200136"),
    ("F5001","NORTHWEST","Surabaya","Ruko Northwest Citraland NV2 No.53, Babat Jerawat, Surabaya","https://maps.app.goo.gl/fg6K2erhzMhvSpVBA","628175200139"),
    ("F5002","SUKOLILO","Surabaya","Eastern Park AB, Jl. Raya Sukolilo Mulia No.A15, Surabaya","https://maps.app.goo.gl/afuqFY3L2G1ow8JRA","628175200137"),
    ("F5003","BUKITMAS","Surabaya","Ruko Wisata Bukitmas 2, Jalan Wisata Bukit Mas Blok RH No.2","https://maps.app.goo.gl/8kavY5TdMcMnb3888","628175200140"),
    ("C5001","CBD CIBUBUR","Bekasi","Ruko CBD Citra Grand Blok ER 2 No 6, Jatisampurna, Bekasi","https://maps.app.goo.gl/MKMTX6h5EwD4VYh19","628175200141"),
    ("C5004","PERMATA CIBUBUR","Bogor","Ruko Permata Cibubur, Jl. Raya Permata Cibubur No.6, Bogor","https://maps.app.goo.gl/JNEUxD692oGLcMiM7","6281944214740"),
    ("C5003","CANADIAN CIBUBUR","Bogor","Kota Wisata Ruko Canadian Broadway, Jl. Raya Kota Wisata, Bogor","https://maps.app.goo.gl/nUubmw5e7fXGTgfNA","6281944220051"),
    ("C4002","TARUMA","Karawang","Ruko Grand Taruma, Jl. Darmawangsa III No.10 D/9, Karawang","https://maps.app.goo.gl/aScJsXL1SAERoPKk9","6287722724809"),
    ("A4001","BUDI RAYA","Jakarta","Jakarta Selatan","",""),
    ("A6001","BENHIL","Jakarta","Jl. Bendungan Hilir No.98, Tanah Abang, Jakarta Pusat","https://maps.app.goo.gl/rHZzRTikB5SHfBXDA",""),
    ("A4003","KEBON JERUK","Jakarta","Jakarta Barat","",""),
    ("A4004","KUNINGAN","Jakarta","Jakarta Selatan","",""),
    ("A5008","MUWARDI","Jakarta","Jl. Muwardi No.42, Grogol, Jakarta Barat","https://maps.app.goo.gl/ER1bq46dGGZTdEqZ6","87722724816"),
    ("B4001","GOLDEN","Tangerang","Tangerang","",""),
    ("B5007","FIORENZA CIKUPA","Tangerang","Ruko Fiorenza, JL. Raya H. Cinde Lakoni No.17/52, Cikupa","https://maps.app.goo.gl/4k6Ke8hLkbxEJaKJA","6281944210513"),
    ("B5004","MILENIAL PIK 2","Tangerang","Rukan Millenial Blok B-07, PIK 2, Teluknaga, Tangerang","https://maps.app.goo.gl/M7sPAxscgdmXaKkA8","6281944219566"),
    ("B5005","TAMAN ASRI","Tangerang","Ruko Taman Asri, Jl. Taman Asri Utama Blok I No.2, Tangerang","https://maps.app.goo.gl/h4W5qt1xxaoff2di6","6287722724806"),
    ("B5006","PALEM SEMI","Tangerang","Komplek Ruko Barcelona No.75 Jl Palem Raja, Karawaci, Tangerang","https://maps.app.goo.gl/LcEGd6ejFeMD3n7U8","628175200134"),
    ("B5008","TAMAN ROYAL","Tangerang","Ruko Taman Royal Permata Niaga 3 No.3, Kec Tangerang","https://maps.app.goo.gl/G55mvXT7ue28tM8B8","6287722724805"),
    ("B4002","PUSTEK","Tangerang Selatan","Jl. Raya Puspitek No.48, Buaran, Serpong, Tangerang Selatan","https://maps.app.goo.gl/PGc77Hcwc1tfPXgX8","6281913580778"),
    ("B5003","BENDA PAMULANG","Tangerang Selatan","Ruko Pamulang Permai No.F1/27, Pamulang, Tangerang Selatan","https://maps.app.goo.gl/v7TbA6F1Xda3SAEG6","6281944215238"),
]

outlets_list = []
for code, name, city, addr, maps, hp in HARDCODED_OUTLETS:
    outlets_list.append({
        'store_code': code,
        'name': name,
        'city': city,
        'address': addr,
        'maps': maps if maps else None,
        'wa': fmt_wa(hp),
    })
print(f"  Outlets: {len(outlets_list)}")

# ── 5. Write JSON ──────────────────────────────────────────────────────────
with open(os.path.join(OUT_DIR, 'products.json'), 'w', encoding='utf-8') as f:
    json.dump(products_list, f, ensure_ascii=False, separators=(',',':'))

with open(os.path.join(OUT_DIR, 'stock.json'), 'w', encoding='utf-8') as f:
    json.dump(stock_index, f, ensure_ascii=False, separators=(',',':'))

with open(os.path.join(OUT_DIR, 'outlets.json'), 'w', encoding='utf-8') as f:
    json.dump(outlets_list, f, ensure_ascii=False, separators=(',',':'))

print("\n✅ Done! Output:")
for fn in ['products.json','stock.json','outlets.json']:
    path = os.path.join(OUT_DIR, fn)
    size = os.path.getsize(path)
    print(f"  data/{fn} — {size/1024:.1f} KB")
