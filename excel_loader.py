import math
import re
import difflib
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

# Known brands list for extraction
KNOWN_BRANDS = [
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
    "JHFOODS", "LIPTON", "ANNY", "SAHIBA", "BAMBOO TREE", "FARMER", "ACECOOK", "WAI WAI THAILAND", "PILLSBURY",
    "PARACHUTE", "MYM", "RENUKA", "YEN NHUNG", "CARNATION", "WABU", "KHANUM", "OYAKATA", "PRIMA", "KAIJAE",
    "HORLICKS", "HERITAGE AFRIKA", "OISHI", "VIMTO", "ML SQUID", "KOH-KAE", "COCK", "DETTOL", "MAMA'S CHOICE",
    "MAE KRUA", "SAGIKO", "SUNFLOWER", "GREEN TABLE", "JONGGA", "VIET NAM", "SHANA", "UPASTRY", "NOODLE HOUSE",
    "TAKIS", "ELEFANT", "HYPER MALT", "HOT CHIP", "SZU SHEN PO", "AGARBATTI", "NARCISSUS", "EVERBEST", "HEALTHY BOY",
    "KINGZEST", "HAIDILAO", "MAO XIONG", "HERBEX", "BAMBOO TREE", "PRESIDENT", "JIADUOBAO", "HIKARI MISO", "LAKOVO",
    "YOPOKKI", "CYPRESSA", "MEHEK", "GINGERBON", "PULMUONE", "WOK FOODS", "BAIJIA", "WEI LIH", "GOLD KILI",
    "PCD", "PRB", "SLINMY", "KAILO", "HUNG PHAT", "SHAN WAI", "HERR'S", "PATAK'S", "POR KWAN", "CARABAO", "WEIJUTE",
    "COFE", "RABBIT", "OTAFUKU", "KHONG DO", "TAMANOI", "JIA BRAND", "SHAN WAI", "YANCO", "HUNG PHAT", "MP", "WANT WANT",
    "YAN LONG", "PAN", "JING YI GEN", "GOGI", "YUANFU BRAND", "MAI WA", "SUKINA", "RAFHAN", "KIMHO", "MOGUMOGU",
    "YUM YUM", "CHIU CHOW", "TARO", "MEGA", "HAOHAO", "JUB JUB", "SKYBIRD", "MEIJI H.PANDA", "ROYAL TIGER", "SAMYANG",
    "COCON", "GENKI RAMUNE", "LAO GAN MA", "IFAD", "LKK", "YAMASA", "JONGGA", "LONGLIFE", "TONGYI", "MAN TANG XIAN",
    "MAE NAPA", "ELEPHANT", "CHUPA CHUPS", "MARUKOME", "EAGLOBE", "SQUID", "MAO XIONG", "EFP", "HEINZ", "BINGGRAE",
    "MINI MELTS", "SICHUAN WANG", "SEMPIO", "NITTAYA", "A", "ATOOM", "HENG SHUN", "JIABAO", "AQUAPEARL", "PERFIT",
    "BRITANNIA", "CROWN", "FLYING GOOSE"
]


class ExcelLoader:
    def __init__(self, vendor_root: Path):
        self.vendor_root = Path(vendor_root)

    @staticmethod
    def _normalize_header(header: str) -> str:
        return re.sub(r"[^a-z0-9]", "_", str(header).strip().lower())

    @staticmethod
    def _split_quantity(quantity: Optional[str]) -> Tuple[Optional[int], Optional[str]]:
        """Split quantity into (moq, weight).
        
        Examples:
        - "10x1kg" -> (10, "1kg")
        - "500g*12" -> (12, "500g")
        - "2 x 5 KG" -> (2, "5 KG")
        
        Returns (None, None) if quantity is None or doesn't match pattern.
        """
        if not quantity:
            return (None, None)
        
        quantity = str(quantity).strip()
        
        # Match common patterns: 10x1kg, 500g*12, 2 x 5 KG, etc.
        # Pattern 1: NUMBER [xX*] NUMBER UNIT (e.g., 10x1kg, 2 x 500 ml)
        match = re.match(r'(\d+)\s*[xX*]\s*(\d+\s*[a-zA-Z]*)', quantity, re.IGNORECASE)
        if match:
            moq = int(match.group(1))
            weight = match.group(2).strip()
            return (moq, weight)
        
        # Pattern 2: NUMBER UNIT [xX*] NUMBER (e.g., 500g*12, 1kg x 10)
        match = re.match(r'(\d+\s*[a-zA-Z]*)\s*[xX*]\s*(\d+)', quantity, re.IGNORECASE)
        if match:
            weight = match.group(1).strip()
            moq = int(match.group(2))
            return (moq, weight)
        
        return (None, None)

    @staticmethod
    def _extract_brand(text: str) -> Optional[str]:
        """Extract brand from text using known brands list.
        
        Returns the brand name if found, None otherwise.
        """
        if not text:
            return None
        
        text = str(text)
        for brand_candidate in KNOWN_BRANDS:
            pattern = r'\b' + re.escape(brand_candidate) + r'\b'
            if re.search(pattern, text, flags=re.IGNORECASE):
                return brand_candidate
        
        return None

    @staticmethod
    def _remove_brand_from_text(text: str, brand: Optional[str]) -> str:
        """Remove brand name from text if present."""
        if not text or not brand:
            return text
        
        pattern = r'\b' + re.escape(brand) + r'\b'
        return re.sub(pattern, '', text, flags=re.IGNORECASE).strip()

    @staticmethod
    def _is_missing(value: Any) -> bool:
        if value is None:
            return True
        if isinstance(value, float) and math.isnan(value):
            return True
        if isinstance(value, str) and value.strip().lower() in ("nan", "none", ""):
            return True
        return False

    def _map_columns(self, columns: List[str]) -> Dict[str, str]:
        mapping = {}
        lower_cols = {self._normalize_header(c): c for c in columns}
        normalized_cols = list(lower_cols.keys())
        field_map = {
            "sku": ["sku", "vendor_sku", "sr_sku", "code", "product_code", "item_code", "artikelnummer", "artikel_nr"],
            "product_name": ["name", "product_name", "description", "item", "item_name"],
            "brand": ["brand", "supplier"],
            "pack": ["pack", "package", "carton", "ctn_pack", "size"],
            "unit": ["unit", "uom", "units"],
            "ctn_qty": ["ctn_qty", "ctn quantity", "carton_qty", "carton quantity", "quantity_ctn", "ctn", "pc/ctn", "pcs_per_ctn", "pack_qty", "pcctn", "ctn ty", "ctn_ty", "ctn-ty"],
            "price": ["price", "unit_price", "purchase_price", "rate", "sale_price"],
            "stock": ["stock", "available", "qty", "quantity", "available_quantity", "in_stock"],
            "barcode": ["barcode", "ean", "upc"],
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
            barcode = row.get(mapping.get("barcode")) if mapping.get("barcode") else None
            
            # Try to extract brand and ctn_qty from product_name if not already present
            extracted_brand = None
            extracted_ctn_qty = None
            extracted_pack_unit = None
            if product_name:
                extracted_brand = self._extract_brand(str(product_name))
                if self._is_missing(brand) and extracted_brand:
                    brand = extracted_brand
                
                # Try to extract quantity pattern from product_name
                quantity_match = re.search(r'(\d+\s*[xX*]\s*\d+\s*\w*|\d+\s*\w+\s*[xX*]\s*\d+)', str(product_name))
                if quantity_match:
                    quantity = quantity_match.group(0).strip()
                    extracted_ctn_qty, extracted_pack_unit = self._split_quantity(quantity)
                    if self._is_missing(ctn_qty) and extracted_ctn_qty is not None:
                        ctn_qty = extracted_ctn_qty
                    if extracted_pack_unit and (self._is_missing(pack) or self._is_missing(unit)):
                        # parse the pack_unit (e.g., "500g" -> pack="500g", unit="g")
                        parsed_pack, parsed_unit = self._parse_pack_unit(extracted_pack_unit)
                        if parsed_pack and self._is_missing(pack):
                            pack = parsed_pack
                        if parsed_unit and self._is_missing(unit):
                            unit = parsed_unit
            
            item = {
                "sku": None if self._is_missing(sku) else str(sku).strip(),
                "product_name": None if self._is_missing(product_name) else str(product_name).strip(),
                "brand": None if self._is_missing(brand) else str(brand).strip(),
                "pack": None if self._is_missing(pack) else str(pack).strip(),
                "unit": None if self._is_missing(unit) else str(unit).strip(),
                "barcode": None if self._is_missing(barcode) else str(barcode).strip(),
                "ctn_qty": int(ctn_qty) if self._is_int_like(ctn_qty) else 0,
                "price": float(price) if self._is_numeric(price) else 0.0,
                "stock": "" if self._is_missing(stock) else str(stock).strip(),
                "last_updated": datetime.utcnow().isoformat(),
                "raw_excel": row.to_dict(),
            }
            # compute sr_sku as vendor prefix + sku when possible
            if sku and vendor:
                item["sr_sku"] = f"{vendor}-{str(sku).strip()}"
            else:
                item["sr_sku"] = None
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
        m = re.search(r"(\d+)\s*[xX]\s*(\d+(?:[\.,]\d+)?)\s*(kg|g|gm|gr|ml|l|ltr|pcs|pc)\b", text, re.IGNORECASE)
        if m:
            n1 = m.group(1)
            n2 = m.group(2).replace(",", ".")
            unit = self._normalize_unit(m.group(3))
            pack = f"{n1}x{n2}{unit}"
            return pack, unit
        # try single quantity unit like 500 g, 250ml, 1kg
        m2 = re.search(r"(\d+(?:[\.,]\d+)?)\s*(kg|g|gm|gr|ml|l|ltr|pcs|pc)\b", text, re.IGNORECASE)
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
