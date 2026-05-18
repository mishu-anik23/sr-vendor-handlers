import re
import difflib
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
        normalized_cols = list(lower_cols.keys())
        field_map = {
            "sku": ["sku", "code", "product_code", "item_code", "ean", "barcode", "artikelnummer", "artikel_nr"],
            "product_name": ["name", "product_name", "description", "item", "item_name"],
            "brand": ["brand", "supplier"],
            "pack": ["pack", "package", "carton", "ctn_pack", "size"],
            "unit": ["unit", "uom", "units"],
            "ctn_qty": ["ctn_qty", "ctn_qty", "ctn quantity", "carton_qty", "carton quantity", "quantity_ctn", "ctn"],
            "price": ["price", "unit_price", "purchase_price", "rate", "sale_price"],
            "stock": ["stock", "available", "qty", "quantity", "available_quantity", "in_stock"],
        }
        for target, names in field_map.items():
            # try exact normalized matches first
            found = False
            for name in names:
                normalized = self._normalize_header(name)
                if normalized in lower_cols:
                    mapping[target] = lower_cols[normalized]
                    found = True
                    break
            if found:
                continue
            # For SKU specifically, prefer any header that contains 'sku' in its normalized form
            if target == "sku":
                candidates = [k for k in normalized_cols if "sku" in k]
                if candidates:
                    # prefer exact 'sku' or 'vendor_sku' if present
                    preferred = None
                    for p in ("sku", "vendor_sku"):
                        if p in candidates:
                            preferred = p
                            break
                    chosen = preferred or candidates[0]
                    mapping[target] = lower_cols[chosen]
                    continue
            # fallback: fuzzy match against headers
            for name in names:
                normalized = self._normalize_header(name)
                matches = difflib.get_close_matches(normalized, normalized_cols, n=1, cutoff=0.7)
                if matches:
                    mapping[target] = lower_cols[matches[0]]
                    found = True
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
            sku = row.get(mapping.get("sku")) if mapping.get("sku") else None
            product_name = row.get(mapping.get("product_name")) if mapping.get("product_name") else None
            brand = row.get(mapping.get("brand")) if mapping.get("brand") else None
            pack = row.get(mapping.get("pack")) if mapping.get("pack") else None
            unit = row.get(mapping.get("unit")) if mapping.get("unit") else None
            ctn_qty = row.get(mapping.get("ctn_qty")) if mapping.get("ctn_qty") else None
            price = row.get(mapping.get("price")) if mapping.get("price") else None
            stock = row.get(mapping.get("stock")) if mapping.get("stock") else None
            item = {
                "sku": str(sku).strip() if sku is not None else None,
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
            # If pack or unit missing, try to parse from product_name
            pack_missing = (not item.get("pack") or str(item.get("pack")).lower() in ("nan", "none", ""))
            unit_missing = (not item.get("unit") or str(item.get("unit")).lower() in ("nan", "none", ""))
            if pack_missing or unit_missing:
                parsed_pack, parsed_unit = self._parse_pack_unit(item.get("product_name") or "")
                if parsed_pack and pack_missing:
                    item["pack"] = parsed_pack
                if parsed_unit and unit_missing:
                    item["unit"] = parsed_unit

            if not item["sku"] and item["product_name"]:
                item["sku"] = item["product_name"]
            if item["product_name"]:
                products.append(item)
        return products

    @staticmethod
    def _normalize_unit(unit: str) -> str:
        if not unit:
            return ""
        u = unit.strip().lower()
        u = u.replace("gm", "g").replace("gr", "g").replace("kgs", "kg").replace("ltr", "l")
        u = u.replace("pcs", "pcs").replace("pc", "pcs")
        return u

    def _parse_pack_unit(self, name: str) -> (str, str):
        """Extract a pack string and unit from a product name.

        Examples matched: '6x330ml', '500 g', '1 kg', '12 pcs', '2 x 400 ml'
        Returns (pack, unit) where either may be empty string if not found.
        """
        if not name:
            return "", ""
        text = str(name)
        # try pattern like 6x330ml or 2 x 400 ml
        m = re.search(r"(\d+)\s*[xX]\s*(\d+(?:[\.,]\d+)?)\s*(kg|g|gm|gr|ml|l|ltr|pcs|pc)\b", text)
        if m:
            n1 = m.group(1)
            n2 = m.group(2).replace(",", ".")
            unit = self._normalize_unit(m.group(3))
            pack = f"{n1}x{n2}{unit}"
            return pack, unit
        # try single quantity unit like 500 g, 250ml, 1kg
        m2 = re.search(r"(\d+(?:[\.,]\d+)?)\s*(kg|g|gm|gr|ml|l|ltr|pcs|pc)\b", text)
        if m2:
            qty = m2.group(1).replace(",", ".")
            unit = self._normalize_unit(m2.group(2))
            pack = f"{qty}{unit}"
            return pack, unit
        return "", ""

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
