import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd


class ExcelLoader:
    def __init__(self, vendor_root: Path):
        self.vendor_root = Path(vendor_root)

    @staticmethod
    def _normalize_header(header: str) -> str:
        return re.sub(r"[^a-z0-9]", "_", str(header).strip().lower())

    def _map_columns(self, columns: List[str]) -> Dict[str, str]:
        mapping = {}
        lower_cols = {self._normalize_header(c): c for c in columns}
        field_map = {
            "source_id": ["sku", "code", "product_code", "item_code", "ean", "barcode"],
            "product_name": ["name", "product_name", "description", "item", "item_name"],
            "brand": ["brand", "supplier"],
            "pack": ["pack", "package", "carton", "ctn_pack", "size"],
            "unit": ["unit", "uom", "units"],
            "ctn_qty": ["ctn_qty", "ctn_qty", "ctn quantity", "carton_qty", "carton quantity", "quantity_ctn", "ctn"],
            "price": ["price", "unit_price", "purchase_price", "rate", "sale_price"],
            "stock": ["stock", "available", "qty", "quantity", "available_quantity", "in_stock"],
        }
        for target, names in field_map.items():
            for name in names:
                normalized = self._normalize_header(name)
                if normalized in lower_cols:
                    mapping[target] = lower_cols[normalized]
                    break
        return mapping

    def find_latest_excel_file(self, vendor: str) -> Optional[Path]:
        vendor_path = self.vendor_root / vendor / "products"
        if not vendor_path.exists() or not vendor_path.is_dir():
            return None
        excel_files = list(vendor_path.glob("*.xlsx"))
        if not excel_files:
            return None
        return max(excel_files, key=lambda p: p.stat().st_mtime)

    def load_products(self, vendor: str) -> List[Dict[str, Any]]:
        excel_path = self.find_latest_excel_file(vendor)
        if not excel_path:
            return []
        df = pd.read_excel(excel_path, engine="openpyxl")
        df = df.where(pd.notnull(df), None)
        mapping = self._map_columns(list(df.columns))
        products = []
        for index, row in df.iterrows():
            source_id = row.get(mapping.get("source_id")) if mapping.get("source_id") else None
            product_name = row.get(mapping.get("product_name")) if mapping.get("product_name") else None
            brand = row.get(mapping.get("brand")) if mapping.get("brand") else None
            pack = row.get(mapping.get("pack")) if mapping.get("pack") else None
            unit = row.get(mapping.get("unit")) if mapping.get("unit") else None
            ctn_qty = row.get(mapping.get("ctn_qty")) if mapping.get("ctn_qty") else None
            price = row.get(mapping.get("price")) if mapping.get("price") else None
            stock = row.get(mapping.get("stock")) if mapping.get("stock") else None
            item = {
                "source_id": str(source_id).strip() if source_id is not None else None,
                "product_name": str(product_name).strip() if product_name is not None else None,
                "brand": str(brand).strip() if brand is not None else None,
                "pack": str(pack).strip() if pack is not None else None,
                "unit": str(unit).strip() if unit is not None else None,
                "ctn_qty": int(ctn_qty) if self._is_int_like(ctn_qty) else 0,
                "price": float(price) if self._is_numeric(price) else 0.0,
                "stock": str(stock).strip() if stock is not None else "",
                "last_updated": datetime.utcnow().isoformat(),
                "raw_excel": row.to_dict(),
            }
            if not item["source_id"] and item["product_name"]:
                item["source_id"] = item["product_name"]
            if item["product_name"]:
                products.append(item)
        return products

    @staticmethod
    def _is_int_like(value: Any) -> bool:
        if value is None:
            return False
        try:
            int(value)
            return True
        except (ValueError, TypeError):
            return False

    @staticmethod
    def _is_numeric(value: Any) -> bool:
        if value is None:
            return False
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False
