# Sunrise Supermarket Vendor Manager

A PyQt application for managing vendor product lists, syncing Excel product sheets into SQLite, creating orders, and tracking billing history.

## Features

- Discover vendor folders dynamically under `data/vendors`
- Load the latest Excel product sheet from each vendor's `products/` folder
- Upsert product rows into per-vendor SQLite tables
- Display product inventory and statistics
- Create order sheets with quantity, brand, and carton quantity
- Save orders into the SQLite database and export a timestamped Excel order file

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
python main.py
```

## Notes

- Order exports are written to `data/vendors/<vendor>/orders/`
- Vendor-specific product tables are stored in `data/vendor_app.db`
- The app uses `logo-sr-tmp.jpeg` as the company logo in the UI
