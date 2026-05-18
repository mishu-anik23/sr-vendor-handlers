from pathlib import Path
import json

from excel_loader import ExcelLoader
from db_manager import DatabaseManager

vendor = 'eurasiamart'
vendor_root = Path('data') / 'vendors'

loader = ExcelLoader(vendor_root)
db = DatabaseManager(Path('data') / 'vendor_app.db')

try:
    products = loader.load_products(vendor)
    print(f'Found {len(products)} products for {vendor}')
    for p in products[:5]:
        print(json.dumps({'sku': p.get('sku'), 'product_name': p.get('product_name')}, ensure_ascii=False))
    db.upsert_vendor_products(vendor, products)
    rows = db.get_vendor_products(vendor)
    print(f'After upsert, {len(rows)} rows in DB (showing 5):')
    for r in rows[:5]:
        print(json.dumps(r, ensure_ascii=False, default=str))
    print('Stats:', db.get_vendor_statistics(vendor))
except Exception as e:
    import traceback
    print('ERROR', e)
    traceback.print_exc()
