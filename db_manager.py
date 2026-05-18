import json
import re
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional


class DatabaseManager:
    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(self.db_path))
        self.conn.row_factory = sqlite3.Row
        self._ensure_order_tables()

    def _execute(self, sql: str, params: tuple = ()) -> sqlite3.Cursor:
        cursor = self.conn.cursor()
        cursor.execute(sql, params)
        return cursor

    @staticmethod
    def sanitize_vendor_name(vendor: str) -> str:
        token = re.sub(r"[^a-zA-Z0-9]+", "_", vendor.strip().lower())
        return token.strip("_") or "vendor"

    def vendor_table_name(self, vendor: str) -> str:
        return f"vendor_{self.sanitize_vendor_name(vendor)}"

    def _ensure_order_tables(self) -> None:
        self._execute(
            """
            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY,
                vendor TEXT NOT NULL,
                created_at TEXT NOT NULL,
                total_amount REAL NOT NULL,
                order_filename TEXT,
                notes TEXT
            )
            """
        )
        self._execute(
            """
            CREATE TABLE IF NOT EXISTS order_items (
                id INTEGER PRIMARY KEY,
                order_id INTEGER NOT NULL,
                source_id TEXT,
                product_name TEXT,
                brand TEXT,
                quantity INTEGER,
                ctn_qty INTEGER,
                unit_price REAL,
                total_price REAL,
                package TEXT,
                raw_json TEXT,
                FOREIGN KEY(order_id) REFERENCES orders(id)
            )
            """
        )
        self.conn.commit()

    def create_vendor_table(self, vendor: str) -> None:
        table_name = self.vendor_table_name(vendor)
        self._execute(
            f"""
            CREATE TABLE IF NOT EXISTS `{table_name}` (
                id INTEGER PRIMARY KEY,
                source_id TEXT UNIQUE,
                product_name TEXT,
                brand TEXT,
                pack TEXT,
                unit TEXT,
                ctn_qty INTEGER,
                price REAL,
                stock TEXT,
                last_updated TEXT,
                extra_json TEXT
            )
            """
        )
        self.conn.commit()

    def upsert_vendor_products(self, vendor: str, products: List[Dict[str, Any]]) -> None:
        self.create_vendor_table(vendor)
        table_name = self.vendor_table_name(vendor)
        cursor = self.conn.cursor()
        for product in products:
            source_id = str(product.get("source_id") or product.get("product_name") or "").strip()
            if not source_id:
                continue
            product_name = str(product.get("product_name") or "").strip()
            brand = str(product.get("brand") or "").strip()
            pack = str(product.get("pack") or "").strip()
            unit = str(product.get("unit") or "").strip()
            ctn_qty = int(product.get("ctn_qty") or 0)
            price = float(product.get("price") or 0.0)
            stock = str(product.get("stock") or "").strip()
            last_updated = product.get("last_updated") or datetime.utcnow().isoformat()
            extra_json = json.dumps(product, default=str)
            cursor.execute(
                f"""
                INSERT INTO `{table_name}`
                    (source_id, product_name, brand, pack, unit, ctn_qty, price, stock, last_updated, extra_json)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(source_id) DO UPDATE SET
                    product_name=excluded.product_name,
                    brand=excluded.brand,
                    pack=excluded.pack,
                    unit=excluded.unit,
                    ctn_qty=excluded.ctn_qty,
                    price=excluded.price,
                    stock=excluded.stock,
                    last_updated=excluded.last_updated,
                    extra_json=excluded.extra_json
                """,
                (source_id, product_name, brand, pack, unit, ctn_qty, price, stock, last_updated, extra_json),
            )
        self.conn.commit()

    def get_vendor_products(self, vendor: str) -> List[Dict[str, Any]]:
        table_name = self.vendor_table_name(vendor)
        self.create_vendor_table(vendor)
        cursor = self._execute(f"SELECT * FROM `{table_name}` ORDER BY product_name COLLATE NOCASE")
        return [dict(row) for row in cursor.fetchall()]

    def get_vendor_statistics(self, vendor: str) -> Dict[str, Any]:
        table_name = self.vendor_table_name(vendor)
        self.create_vendor_table(vendor)
        cursor = self._execute(
            f"SELECT COUNT(*) AS product_count, SUM(ctn_qty * price) AS inventory_value FROM `{table_name}`"
        )
        row = cursor.fetchone()
        return {
            "vendor": vendor,
            "product_count": int(row["product_count"] or 0),
            "inventory_value": float(row["inventory_value"] or 0.0),
        }

    def get_all_vendor_statistics(self, vendors: List[str]) -> List[Dict[str, Any]]:
        return [self.get_vendor_statistics(v) for v in vendors]

    def save_order(
        self,
        vendor: str,
        items: List[Dict[str, Any]],
        total_amount: float,
        order_filename: str,
        notes: Optional[str] = None,
    ) -> int:
        now = datetime.utcnow().isoformat()
        cursor = self.conn.cursor()
        cursor.execute(
            "INSERT INTO orders (vendor, created_at, total_amount, order_filename, notes) VALUES (?, ?, ?, ?, ?)",
            (vendor, now, total_amount, order_filename, notes or ""),
        )
        order_id = cursor.lastrowid
        for item in items:
            source_id = str(item.get("source_id") or item.get("product_name") or "").strip()
            product_name = str(item.get("product_name") or "").strip()
            brand = str(item.get("brand") or "").strip()
            quantity = int(item.get("quantity") or 0)
            ctn_qty = int(item.get("ctn_qty") or 0)
            unit_price = float(item.get("unit_price") or 0.0)
            total_price = float(item.get("total_price") or 0.0)
            package = str(item.get("pack") or "").strip()
            raw_json = json.dumps(item, default=str)
            cursor.execute(
                "INSERT INTO order_items (order_id, source_id, product_name, brand, quantity, ctn_qty, unit_price, total_price, package, raw_json) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (order_id, source_id, product_name, brand, quantity, ctn_qty, unit_price, total_price, package, raw_json),
            )
        self.conn.commit()
        return order_id

    def get_orders(self, vendor: Optional[str] = None) -> List[Dict[str, Any]]:
        if vendor:
            cursor = self._execute("SELECT * FROM orders WHERE vendor = ? ORDER BY created_at DESC", (vendor,))
        else:
            cursor = self._execute("SELECT * FROM orders ORDER BY created_at DESC")
        return [dict(row) for row in cursor.fetchall()]

    def get_order_items(self, order_id: int) -> List[Dict[str, Any]]:
        cursor = self._execute("SELECT * FROM order_items WHERE order_id = ?", (order_id,))
        return [dict(row) for row in cursor.fetchall()]
