# Sunrise Supermarket Vendor Manager

Sunrise Supermarket Vendor Manager is a desktop PyQt5 application (with an optional mobile web UI) to
ingest vendor Excel product lists, maintain per-vendor product catalogs in SQLite, create and manage
orders, export order sheets (Excel/PDF), and operate a simple mobile inventory interface with barcode
scanning.

**Key Features**
- **Vendor sync:** Discover vendor folders under `data/vendors` and load the latest Excel product file from each vendor's `products/` folder using `ExcelLoader`.
- **Smart Excel parsing:** `excel_loader.py` performs header normalization, fuzzy column mapping, brand extraction from product names (uses a built-in `KNOWN_BRANDS` list), pack/unit parsing (e.g., `6x330ml`, `500 g`), and quantity splitting. It returns normalized product dicts ready for DB upsert.
- **Per-vendor database tables:** `db_manager.py` creates and manages per-vendor tables (safe names like `vendor_<name>`), runs migrations, consolidates SKUs from `extra_json`, and provides statistics and order persistence.
- **Order management:** Create orders from the GUI, save to SQLite (`orders` + `order_items`), export timestamped Excel files, and export well-formatted PDFs (uses `reportlab`).
- **Product inventory UI:** Inventory list, search, sort, brand filter, export per-vendor / all-vendors Excel exports, and product editing dialogs.
- **Vendor profiles:** Store business details (IBAN, tax IDs, contact, website, etc.) and include them in PDF exports and mobile view.
- **Mobile web UI:** Lightweight Flask server in `web_server.py` exposes an inventory browser and editor with barcode scanning using `html5-qrcode`. Start automatically from the desktop app if Flask is installed.
- **Utilities & resilience:** Robust handling of missing columns, legacy fields (`source_id`), SKU consolidation, and fixes for database index issues.

**Installation**

1. Create a Python environment (recommended):

```bash
python -m venv .venv
source .venv/Scripts/activate  # Windows: .venv\Scripts\activate
```
2. Install dependencies:

```bash
pip install -r requirements.txt
```

Optional extras:
- `reportlab` — for PDF export (used by the app). Install with `pip install reportlab`.
- `cryptography` — to enable HTTPS for the mobile Flask server.

**Run (desktop GUI)**

```bash
python main.py
```

The GUI provides buttons to:
- Sync selected vendor / Sync all vendors (`Sync selected vendor`, `Sync all vendors`)
- Reset stored products for a vendor (`Reset selected vendor products`)
- Open product inventory (`Product inventory`) and edit items
- Create and edit orders (`Create order`) and export to Excel/PDF
- Open vendor business details (`Vendor details`)

When the app starts it will attempt to start the mobile server (if Flask is installed) and show the server URL in the UI.

**Run (mobile / web UI)**

The mobile web UI is implemented in `web_server.py`. It exposes:
- Inventory JSON endpoints (`/api/inventory/<vendor>`) and an editor POST endpoint (`/api/inventory/<vendor>/<sku>`).
- A simple single-page HTML app with barcode scanning (uses `html5-qrcode`) at `/`.

To run manually (if not started by the GUI):

```bash
python -c "from web_server import start_web_server; from db_manager import DatabaseManager; start_web_server(DatabaseManager('data/vendor_app.db'), lambda: [p.name for p in Path('data/vendors').iterdir() if p.is_dir()])"
```

If `cryptography` is installed, the server will try to use HTTPS (self-signed adhoc cert). Otherwise it falls back to HTTP and the GUI will prompt to install `cryptography` for camera support.

**Project layout**
- `main.py` — PyQt5 desktop application and main program logic.
- `excel_loader.py` — Excel parsing, header mapping, brand extraction, pack/unit parsing.
- `db_manager.py` — SQLite-backed persistence for vendor tables, orders, order_items and vendor_profiles.
- `web_server.py` — Flask-based mobile inventory UI with barcode scanning and edit endpoint.
- `tests/sync_test.py` — Small test script demonstrating loading, upserting and DB statistics.
- `data/` — Root for vendor folders and `vendor_app.db` (SQLite DB stored at `data/vendor_app.db`).

**Data directories**
- Vendor products: `data/vendors/<vendor>/products/` (place vendor Excel `.xlsx` files here)
- Vendor orders exports: `data/vendors/<vendor>/orders/` (created when orders are exported)
- Vendor exports: `data/vendors/<vendor>/exports/` or `data/vendors/exports/` for all-vendor exports

**Testing**

Quick local test (prints product counts and DB stats):

```bash
python tests/sync_test.py
```

**Notes & troubleshooting**
- Excel parsing is tolerant to many header naming variations; if a column is not recognized, the loader attempts fuzzy matching and also extracts data from product names.
- If PDF export fails, install `reportlab`.
- If camera access for barcode scanning fails, ensure `cryptography` is installed and accept browser security warnings for the local IP when using HTTPS, or use HTTP.
- The DB manager includes migration helpers to handle older schema differences (e.g., `source_id` -> `sku`) and to avoid problematic UNIQUE constraints on nullable `sr_sku` values.

**Contributing**
- Add vendor Excel files under `data/vendors/<vendor>/products/` and use the GUI `Sync` buttons to import.
- For improvements to brand detection or parsing, update `excel_loader.py` (the `KNOWN_BRANDS` list and parsing helpers).

**Files changed / referenced**
- `main.py`, `excel_loader.py`, `db_manager.py`, `web_server.py`, `tests/sync_test.py`

---

If you want, I can also:
- Run the test script and show output, or
- Add a short usage GIF or screenshots to the README.
