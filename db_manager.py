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
        self.conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._ensure_order_tables()
        self._ensure_vendor_profiles_table()
        self._fix_sr_sku_unique_indexes()  # Fix the incorrect UNIQUE constraint on sr_sku

    def _fix_sr_sku_unique_indexes(self) -> None:
        """Drop UNIQUE indexes on sr_sku for all vendor tables since sr_sku can be NULL
        and NULL values cause issues with UNIQUE constraints in SQLite."""
        cursor = self.conn.cursor()
        # Get all vendor tables
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'vendor_%'")
        tables = [row[0] for row in cursor.fetchall()]
        
        for table_name in tables:
            # Check if the incorrect UNIQUE index exists
            index_name = f"idx_{table_name}_sr_sku"
            cursor.execute("SELECT name FROM sqlite_master WHERE type='index' AND name = ?", (index_name,))
            if cursor.fetchone():
                try:
                    cursor.execute(f"DROP INDEX `{index_name}`")
                    self.conn.commit()
                except sqlite3.OperationalError:
                    pass  # Index might not exist or already dropped

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
                updated_at TEXT,
                total_amount REAL NOT NULL,
                order_filename TEXT,
                notes TEXT
            )
            """
        )
        self._ensure_column_on_orders("updated_at")
        self._execute(
            """
            CREATE TABLE IF NOT EXISTS order_items (
                id INTEGER PRIMARY KEY,
                order_id INTEGER NOT NULL,
                sku TEXT,
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
        # ensure sku column exists on order_items for older DBs
        self._ensure_column_on_order_items("sku")
        # migrate any legacy source_id values to sku on order_items
        self._migrate_order_items_source_id()

    def _ensure_vendor_profiles_table(self) -> None:
        self._execute(
            """
            CREATE TABLE IF NOT EXISTS vendor_profiles (
                vendor TEXT PRIMARY KEY,
                display_name TEXT,
                legal_name TEXT,
                address TEXT,
                country_of_origin TEXT,
                contact_name TEXT,
                contact_email TEXT,
                contact_phone TEXT,
                tax_id TEXT,
                vat_id TEXT,
                iban TEXT,
                bank_name TEXT,
                swift_bic TEXT,
                customer_number TEXT,
                customer_website TEXT,
                extra_json TEXT,
                updated_at TEXT
            )
            """
        )
        self.conn.commit()
        self._ensure_column_on_vendor_profiles("customer_number")
        self._ensure_column_on_vendor_profiles("customer_website")

    def create_vendor_table(self, vendor: str) -> None:
        table_name = self.vendor_table_name(vendor)
        self._execute(
            f"""
            CREATE TABLE IF NOT EXISTS `{table_name}` (
                id INTEGER PRIMARY KEY,
                sku TEXT,
                product_name TEXT,
                brand TEXT,
                pack TEXT,
                unit TEXT,
                ctn_qty INTEGER,
                price REAL,
                stock TEXT,
                barcode TEXT,
                sr_sku TEXT,
                last_updated TEXT,
                extra_json TEXT
            )
            """
        )
        self.conn.commit()
        # ensure sku column and unique index exist for this vendor table (migrate older tables)
        self._ensure_column_on_table(table_name, "sku")
        self._ensure_unique_index(table_name, "sku")
        # ensure barcode and sr_sku columns/index
        self._ensure_column_on_table(table_name, "barcode")
        self._ensure_column_on_table(table_name, "sr_sku")
        # Note: sr_sku is NOT uniquely indexed because many products may not have a vendor-provided SKU (NULL values)
        # and SQLite's UNIQUE constraint treats multiple NULLs as duplicates in some configurations
        # migrate any legacy source_id values to sku for this vendor table
        self._migrate_source_id_to_sku(table_name)
        # consolidate SKUs from extra_json/raw_excel when possible
        self._consolidate_skus(table_name)

    def upsert_vendor_products(self, vendor: str, products: List[Dict[str, Any]]) -> None:
        self.create_vendor_table(vendor)
        table_name = self.vendor_table_name(vendor)
        cursor = self.conn.cursor()
        for product in products:
            sku = str(product.get("sku") or product.get("source_id") or product.get("product_name") or "").strip()
            if not sku:
                continue
            product_name = str(product.get("product_name") or "").strip()
            brand = str(product.get("brand") or "").strip()
            pack = str(product.get("pack") or "").strip()
            unit = str(product.get("unit") or "").strip()
            ctn_qty = int(product.get("ctn_qty") or 0)
            price = float(product.get("price") or 0.0)
            stock = str(product.get("stock") or "").strip()
            barcode = str(product.get("barcode") or "").strip()
            sr_sku = str(product.get("sr_sku") or "").strip()
            last_updated = product.get("last_updated") or datetime.utcnow().isoformat()
            extra_json = json.dumps(product, default=str)
            try:
                cursor.execute(
                    f"""
                    INSERT INTO `{table_name}`
                        (sku, product_name, brand, pack, unit, ctn_qty, price, stock, barcode, sr_sku, last_updated, extra_json)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(sku) DO UPDATE SET
                        product_name=excluded.product_name,
                        brand=excluded.brand,
                        pack=excluded.pack,
                        unit=excluded.unit,
                        ctn_qty=excluded.ctn_qty,
                        price=excluded.price,
                        stock=excluded.stock,
                        barcode=excluded.barcode,
                        sr_sku=excluded.sr_sku,
                        last_updated=excluded.last_updated,
                        extra_json=excluded.extra_json
                    """,
                    (sku, product_name, brand, pack, unit, ctn_qty, price, stock, barcode, sr_sku, last_updated, extra_json),
                )
            except sqlite3.IntegrityError:
                # Fallback for edge cases: perform explicit UPDATE
                cursor.execute(
                    f"UPDATE `{table_name}` SET product_name=?, brand=?, pack=?, unit=?, ctn_qty=?, price=?, stock=?, last_updated=?, extra_json=? WHERE sku=?",
                    (product_name, brand, pack, unit, ctn_qty, price, stock, last_updated, extra_json, sku),
                )
        self.conn.commit()

    def get_vendor_products(self, vendor: str) -> List[Dict[str, Any]]:
        table_name = self.vendor_table_name(vendor)
        self.create_vendor_table(vendor)
        cursor = self._execute(f"SELECT * FROM `{table_name}` ORDER BY product_name COLLATE NOCASE")
        return [dict(row) for row in cursor.fetchall()]

    def update_vendor_product(self, vendor: str, sku: str, update_data: Dict[str, Any]) -> None:
        table_name = self.vendor_table_name(vendor)
        cursor = self.conn.cursor()
        
        fields = []
        values = []
        for k, v in update_data.items():
            fields.append(f"{k} = ?")
            values.append(v)
            
        values.append(sku)
        
        query = f"UPDATE `{table_name}` SET {', '.join(fields)} WHERE sku = ?"
        cursor.execute(query, values)
        self.conn.commit()

    def get_vendor_statistics(self, vendor: str) -> Dict[str, Any]:
        table_name = self.vendor_table_name(vendor)
        self.create_vendor_table(vendor)
        cursor = self._execute(
            f"SELECT COUNT(*) AS product_count, SUM(ctn_qty * price) AS inventory_value, SUM(CASE WHEN barcode IS NOT NULL AND barcode != '' THEN 1 ELSE 0 END) AS barcode_count, COUNT(DISTINCT sr_sku) AS srsku_count FROM `{table_name}`"
        )
        row = cursor.fetchone()
        return {
            "vendor": vendor,
            "product_count": int(row["product_count"] or 0),
            "inventory_value": float(row["inventory_value"] or 0.0),
            "barcode_count": int(row["barcode_count"] or 0),
            "srsku_count": int(row["srsku_count"] or 0),
        }

    def reset_vendor_products(self, vendor: str) -> None:
        table_name = self.vendor_table_name(vendor)
        self.create_vendor_table(vendor)
        self._execute(f"DELETE FROM `{table_name}`")
        self.conn.commit()

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
            sku = str(item.get("sku") or item.get("source_id") or item.get("product_name") or "").strip()
            product_name = str(item.get("product_name") or "").strip()
            brand = str(item.get("brand") or "").strip()
            quantity = int(item.get("quantity") or 0)
            ctn_qty = int(item.get("ctn_qty") or 0)
            unit_price = float(item.get("unit_price") or 0.0)
            total_price = float(item.get("total_price") or 0.0)
            package = str(item.get("pack") or "").strip()
            raw_json = json.dumps(item, default=str)
            cursor.execute(
                "INSERT INTO order_items (order_id, sku, product_name, brand, quantity, ctn_qty, unit_price, total_price, package, raw_json) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (order_id, sku, product_name, brand, quantity, ctn_qty, unit_price, total_price, package, raw_json),
            )
        self.conn.commit()
        return order_id

    def _ensure_column_on_table(self, table_name: str, column: str) -> None:
        # check if column exists, if not add it
        cursor = self.conn.cursor()
        cursor.execute(f"PRAGMA table_info(`{table_name}`)")
        cols = [r[1] for r in cursor.fetchall()]
        if column not in cols:
            cursor.execute(f"ALTER TABLE `{table_name}` ADD COLUMN {column} TEXT")
            self.conn.commit()

    def _ensure_unique_index(self, table_name: str, column: str) -> None:
        # create a unique index if it doesn't exist (used for ON CONFLICT upserts)
        index_name = f"idx_{table_name}_{column}"
        self._execute(f"CREATE UNIQUE INDEX IF NOT EXISTS `{index_name}` ON `{table_name}`(`{column}`)")
        self.conn.commit()

    def _ensure_column_on_order_items(self, column: str) -> None:
        cursor = self.conn.cursor()
        cursor.execute("PRAGMA table_info(order_items)")
        cols = [r[1] for r in cursor.fetchall()]
        if column not in cols:
            cursor.execute(f"ALTER TABLE order_items ADD COLUMN {column} TEXT")
            self.conn.commit()

    def _ensure_column_on_orders(self, column: str) -> None:
        cursor = self.conn.cursor()
        cursor.execute("PRAGMA table_info(orders)")
        cols = [r[1] for r in cursor.fetchall()]
        if column not in cols:
            cursor.execute(f"ALTER TABLE orders ADD COLUMN {column} TEXT")
            self.conn.commit()

    def _ensure_column_on_vendor_profiles(self, column: str) -> None:
        cursor = self.conn.cursor()
        cursor.execute("PRAGMA table_info(vendor_profiles)")
        cols = [r[1] for r in cursor.fetchall()]
        if column not in cols:
            cursor.execute(f"ALTER TABLE vendor_profiles ADD COLUMN {column} TEXT")
            self.conn.commit()

    def _migrate_source_id_to_sku(self, table_name: str) -> None:
        # If the legacy `source_id` column exists, copy its values into `sku` where missing.
        cursor = self.conn.cursor()
        cursor.execute(f"PRAGMA table_info(`{table_name}`)")
        cols = [r[1] for r in cursor.fetchall()]
        if "source_id" in cols and "sku" in cols:
            cursor.execute(f"SELECT id, source_id FROM `{table_name}` WHERE (sku IS NULL OR sku = '') AND (source_id IS NOT NULL AND source_id != '')")
            rows = cursor.fetchall()
            for r in rows:
                row_id = r[0]
                source_id = r[1]
                if not source_id:
                    continue
                cursor.execute(f"SELECT COUNT(1) FROM `{table_name}` WHERE sku = ?", (source_id,))
                count = cursor.fetchone()[0]
                if count == 0:
                    cursor.execute(f"UPDATE `{table_name}` SET sku = ? WHERE id = ?", (source_id, row_id))
                else:
                    # make a unique fallback sku using row id
                    fallback = f"{source_id}_{row_id}"
                    cursor.execute(f"UPDATE `{table_name}` SET sku = ? WHERE id = ?", (fallback, row_id))
            self.conn.commit()

    def _migrate_order_items_source_id(self) -> None:
        cursor = self.conn.cursor()
        cursor.execute("PRAGMA table_info(order_items)")
        cols = [r[1] for r in cursor.fetchall()]
        if "source_id" in cols and "sku" in cols:
            cursor.execute("SELECT id, source_id FROM order_items WHERE (sku IS NULL OR sku = '') AND (source_id IS NOT NULL AND source_id != '')")
            rows = cursor.fetchall()
            for r in rows:
                row_id = r[0]
                source_id = r[1]
                if not source_id:
                    continue
                cursor.execute("SELECT COUNT(1) FROM order_items WHERE sku = ?", (source_id,))
                count = cursor.fetchone()[0]
                if count == 0:
                    cursor.execute("UPDATE order_items SET sku = ? WHERE id = ?", (source_id, row_id))
                else:
                    fallback = f"{source_id}_{row_id}"
                    cursor.execute("UPDATE order_items SET sku = ? WHERE id = ?", (fallback, row_id))
            self.conn.commit()

    def _consolidate_skus(self, table_name: str) -> None:
        """Look into `extra_json` to find vendor-provided SKU values and update rows
        where `sku` is empty or equals the product_name (fallback). This helps
        consolidate older rows that used product_name as SKU.
        """
        cursor = self.conn.cursor()
        cursor.execute(f"SELECT id, sku, product_name, extra_json FROM `{table_name}`")
        rows = cursor.fetchall()
        for r in rows:
            row_id = r[0]
            sku = r[1]
            product_name = r[2]
            extra = r[3]
            needs_update = False
            if not sku or (product_name and str(sku).strip() == str(product_name).strip()):
                # try to extract any sku-like field from extra_json
                try:
                    data = json.loads(extra or "{}")
                except Exception:
                    data = {}
                candidate = None
                # check top-level keys in extra/raw_excel for anything with 'sku' in the key
                for k, v in data.items():
                    if "sku" in str(k).lower() and v:
                        candidate = str(v).strip()
                        break
                # also check nested raw_excel if present
                if not candidate and isinstance(data.get("raw_excel"), dict):
                    for k, v in data.get("raw_excel").items():
                        if "sku" in str(k).lower() and v:
                            candidate = str(v).strip()
                            break
                if candidate:
                    # ensure uniqueness
                    cursor.execute(f"SELECT COUNT(1) FROM `{table_name}` WHERE sku = ?", (candidate,))
                    if cursor.fetchone()[0] == 0:
                        cursor.execute(f"UPDATE `{table_name}` SET sku = ? WHERE id = ?", (candidate, row_id))
                        needs_update = True
                    else:
                        # make a unique form
                        fallback = f"{candidate}_{row_id}"
                        cursor.execute(f"UPDATE `{table_name}` SET sku = ? WHERE id = ?", (fallback, row_id))
                        needs_update = True
            if needs_update:
                self.conn.commit()

    def get_vendor_profile(self, vendor: str) -> Dict[str, Any]:
        cursor = self._execute("SELECT * FROM vendor_profiles WHERE vendor = ?", (vendor,))
        row = cursor.fetchone()
        return dict(row) if row else {}

    def save_vendor_profile(self, vendor: str, profile: Dict[str, Any]) -> None:
        now = datetime.utcnow().isoformat()
        display_name = str(profile.get("display_name") or "").strip()
        legal_name = str(profile.get("legal_name") or "").strip()
        address = str(profile.get("address") or "").strip()
        country_of_origin = str(profile.get("country_of_origin") or "").strip()
        contact_name = str(profile.get("contact_name") or "").strip()
        contact_email = str(profile.get("contact_email") or "").strip()
        contact_phone = str(profile.get("contact_phone") or "").strip()
        tax_id = str(profile.get("tax_id") or "").strip()
        vat_id = str(profile.get("vat_id") or "").strip()
        iban = str(profile.get("iban") or "").strip()
        bank_name = str(profile.get("bank_name") or "").strip()
        swift_bic = str(profile.get("swift_bic") or "").strip()
        customer_number = str(profile.get("customer_number") or "").strip()
        customer_website = str(profile.get("customer_website") or "").strip()
        extra_json = json.dumps(profile, default=str)
        self._execute(
            """
            INSERT INTO vendor_profiles
                (vendor, display_name, legal_name, address, country_of_origin, contact_name, contact_email, contact_phone, tax_id, vat_id, iban, bank_name, swift_bic, customer_number, customer_website, extra_json, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(vendor) DO UPDATE SET
                display_name = excluded.display_name,
                legal_name = excluded.legal_name,
                address = excluded.address,
                country_of_origin = excluded.country_of_origin,
                contact_name = excluded.contact_name,
                contact_email = excluded.contact_email,
                contact_phone = excluded.contact_phone,
                tax_id = excluded.tax_id,
                vat_id = excluded.vat_id,
                iban = excluded.iban,
                bank_name = excluded.bank_name,
                swift_bic = excluded.swift_bic,
                customer_number = excluded.customer_number,
                customer_website = excluded.customer_website,
                extra_json = excluded.extra_json,
                updated_at = excluded.updated_at
            """,
            (
                vendor,
                display_name,
                legal_name,
                address,
                country_of_origin,
                contact_name,
                contact_email,
                contact_phone,
                tax_id,
                vat_id,
                iban,
                bank_name,
                swift_bic,
                customer_number,
                customer_website,
                extra_json,
                now,
            ),
        )
        self.conn.commit()

    def get_orders(self, vendor: Optional[str] = None) -> List[Dict[str, Any]]:
        if vendor:
            cursor = self._execute("SELECT * FROM orders WHERE vendor = ? ORDER BY created_at DESC", (vendor,))
        else:
            cursor = self._execute("SELECT * FROM orders ORDER BY created_at DESC")
        return [dict(row) for row in cursor.fetchall()]

    def get_order_items(self, order_id: int) -> List[Dict[str, Any]]:
        cursor = self._execute("SELECT * FROM order_items WHERE order_id = ?", (order_id,))
        return [dict(row) for row in cursor.fetchall()]

    def update_order(
        self,
        order_id: int,
        items: List[Dict[str, Any]],
        total_amount: float,
        notes: Optional[str] = None,
        order_filename: Optional[str] = None,
    ) -> None:
        """Update an existing order and replace its items with the provided list."""
        now = datetime.utcnow().isoformat()
        cursor = self.conn.cursor()
        # Update orders metadata with updated_at timestamp
        if order_filename is not None:
            cursor.execute(
                "UPDATE orders SET total_amount = ?, order_filename = ?, notes = ?, updated_at = ? WHERE id = ?",
                (total_amount, order_filename, notes or "", now, order_id),
            )
        else:
            cursor.execute(
                "UPDATE orders SET total_amount = ?, notes = ?, updated_at = ? WHERE id = ?",
                (total_amount, notes or "", now, order_id),
            )

        # Remove existing items and insert new ones
        cursor.execute("DELETE FROM order_items WHERE order_id = ?", (order_id,))
        for item in items:
            sku = str(item.get("sku") or item.get("source_id") or item.get("product_name") or "").strip()
            product_name = str(item.get("product_name") or "").strip()
            brand = str(item.get("brand") or "").strip()
            quantity = int(item.get("quantity") or 0)
            ctn_qty = int(item.get("ctn_qty") or 0)
            unit_price = float(item.get("unit_price") or 0.0)
            total_price = float(item.get("total_price") or 0.0)
            package = str(item.get("pack") or "").strip()
            raw_json = json.dumps(item, default=str)
            cursor.execute(
                "INSERT INTO order_items (order_id, sku, product_name, brand, quantity, ctn_qty, unit_price, total_price, package, raw_json) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (order_id, sku, product_name, brand, quantity, ctn_qty, unit_price, total_price, package, raw_json),
            )
        self.conn.commit()

    def delete_order(self, order_id: int) -> None:
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM order_items WHERE order_id = ?", (order_id,))
        cursor.execute("DELETE FROM orders WHERE id = ?", (order_id,))
        self.conn.commit()
