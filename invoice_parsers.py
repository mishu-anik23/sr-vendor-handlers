import sys
from pathlib import Path
from typing import Dict, List, Tuple, Any

class BaseInvoiceParser:
    """Base class defining the interface for all invoice parsers."""
    def parse(self, pdf_path: Path) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
        """
        Parses the PDF invoice.
        Returns:
            Tuple of (metadata_dict, rows_list_of_dicts)
        """
        raise NotImplementedError

    def write_excel(self, meta: Dict[str, Any], rows: List[Dict[str, Any]], out_path: Path) -> None:
        """
        Writes the parsed metadata and rows to an Excel sheet.
        """
        raise NotImplementedError


class GFTInvoiceParser(BaseInvoiceParser):
    def parse(self, pdf_path: Path) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
        import pdf_invoice_to_excel as gft
        return gft.parse_pdf(str(pdf_path))
        
    def write_excel(self, meta: Dict[str, Any], rows: List[Dict[str, Any]], out_path: Path) -> None:
        import pdf_invoice_to_excel as gft
        gft.write_excel(meta, rows, str(out_path))


class UnidexInvoiceParser(BaseInvoiceParser):
    def parse(self, pdf_path: Path) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
        import pdf_invoice_to_excel_unidex as unidex
        return unidex.parse_pdf(str(pdf_path))
        
    def write_excel(self, meta: Dict[str, Any], rows: List[Dict[str, Any]], out_path: Path) -> None:
        import pdf_invoice_to_excel_unidex as unidex
        unidex.write_excel(meta, rows, str(out_path))


class AsiaExpressInvoiceParser(BaseInvoiceParser):
    def parse(self, pdf_path: Path) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
        import pdf_to_excel_asiaexpress as ae
        return ae.parse_pdf(str(pdf_path))
        
    def write_excel(self, meta: Dict[str, Any], rows: List[Dict[str, Any]], out_path: Path) -> None:
        import pdf_to_excel_asiaexpress as ae
        ae.write_excel(meta, rows, str(out_path))


class IFTInvoiceParser(BaseInvoiceParser):
    def parse(self, pdf_path: Path) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
        import pdf_invoice_to_excel_ift as ift
        return ift.parse_pdf(str(pdf_path))
        
    def write_excel(self, meta: Dict[str, Any], rows: List[Dict[str, Any]], out_path: Path) -> None:
        import pdf_invoice_to_excel_ift as ift
        ift.write_excel(meta, rows, str(out_path))


class EurFrozenInvoiceParser(BaseInvoiceParser):
    def parse(self, pdf_path: Path) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
        import pdf_invoice_to_excel_eurfrozen as eurfrozen
        return eurfrozen.parse_pdf(str(pdf_path))
        
    def write_excel(self, meta: Dict[str, Any], rows: List[Dict[str, Any]], out_path: Path) -> None:
        import pdf_invoice_to_excel_eurfrozen as eurfrozen
        eurfrozen.write_excel(meta, rows, str(out_path))


class CKInvoiceParser(BaseInvoiceParser):
    def parse(self, pdf_path: Path) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
        import pdf_invoice_to_excel_ck as ck
        return ck.parse_pdf(str(pdf_path))
        
    def write_excel(self, meta: Dict[str, Any], rows: List[Dict[str, Any]], out_path: Path) -> None:
        import pdf_invoice_to_excel_ck as ck
        ck.write_excel(meta, rows, str(out_path))


# Registry mapping vendor directory names to parser instances
PARSERS: Dict[str, BaseInvoiceParser] = {
    "gft": GFTInvoiceParser(),
    "unidex": UnidexInvoiceParser(),
    "asiaexpress": AsiaExpressInvoiceParser(),
    "ift": IFTInvoiceParser(),
    "eurfrozen": EurFrozenInvoiceParser(),
    "ck": CKInvoiceParser(),
}

def get_parser(vendor_name: str) -> BaseInvoiceParser:
    """Returns the parser registered for the given vendor name, or None if not found."""
    return PARSERS.get(vendor_name.lower())
