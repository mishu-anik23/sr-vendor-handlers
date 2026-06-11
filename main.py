import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget, QHBoxLayout, QSpinBox,
)

from db_manager import DatabaseManager
from excel_loader import ExcelLoader

COMPANY_NAME = "Sunrise Supermarket"
COMPANY_ADDRESS = "Schwarzwald Straße 27, 60528 Frankfurt"
LOGO_FILE = Path(__file__).resolve().parent / "logo-sr-tmp.jpeg"
DB_FILE = Path(__file__).resolve().parent / "data" / "vendor_app.db"
VENDOR_ROOT = Path(__file__).resolve().parent / "data" / "vendors"


class OrderDialog(QDialog):
    def __init__(self, vendor: str, products: List[Dict], db: DatabaseManager, parent=None, order_id: Optional[int] = None, initial_items: Optional[List[Dict]] = None):
        super().__init__(parent)
        self.vendor = vendor
        self.products = products
        self.db = db
        self.cart_items: List[Dict] = []  # Store items added to cart
        self.order_id: Optional[int] = order_id
        if initial_items:
            # ensure we copy items so dialog modifications don't mutate caller data
            self.cart_items = [dict(i) for i in initial_items]
        self.selected_product_row: Optional[int] = None
        self.setWindowTitle(f"Create Order - {vendor}")
        self.setMinimumSize(1200, 800)
        self._build_ui()
        self._load_products_table()
        self.pdf_button.setEnabled(bool(self.order_id))
        if self.cart_items:
            self._update_cart_table()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        
        # Header
        header = QHBoxLayout()
        header.addWidget(QLabel(f"<b>Vendor:</b> {self.vendor}"))
        header.addStretch()
        layout.addLayout(header)
        
        # Search and filter panel
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Search Product:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Type product name...")
        self.search_input.textChanged.connect(self._filter_products)
        filter_layout.addWidget(self.search_input)
        
        filter_layout.addWidget(QLabel("Filter by Brand:"))
        self.brand_filter = QComboBox()
        self.brand_filter.addItem("All brands")
        self._populate_brand_filter()
        self.brand_filter.currentTextChanged.connect(self._filter_products)
        filter_layout.addWidget(self.brand_filter)
        layout.addLayout(filter_layout)
        
        # Products table
        layout.addWidget(QLabel("<b>Available Products</b>"))
        self.products_table = QTableWidget(0, 10)
        self.products_table.setHorizontalHeaderLabels(
            ["SKU", "Brand", "Product", "Pack", "Unit", "CTN Qty", "Price", "Qty (0-10)", "Custom Qty", "Add to Cart"]
        )
        header = self.products_table.horizontalHeader()
        for col in range(self.products_table.columnCount()):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        self.products_table.setColumnWidth(0, 140)
        self.products_table.setColumnWidth(1, 120)
        self.products_table.setColumnWidth(2, 420)
        self.products_table.setColumnWidth(3, 120)
        self.products_table.setColumnWidth(4, 80)
        self.products_table.setColumnWidth(5, 90)
        self.products_table.setColumnWidth(6, 100)
        self.products_table.setColumnWidth(7, 90)
        self.products_table.setColumnWidth(8, 100)
        self.products_table.setColumnWidth(9, 100)
        self.products_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.products_table.setSelectionMode(QTableWidget.SingleSelection)
        self.products_table.itemSelectionChanged.connect(self._update_selected_row)
        layout.addWidget(self.products_table)

        add_selected_layout = QHBoxLayout()
        add_selected_layout.addStretch()
        self.add_selected_button = QPushButton("Add selected row to cart")
        self.add_selected_button.clicked.connect(self.on_add_selected_row_to_cart)
        add_selected_layout.addWidget(self.add_selected_button)
        layout.addLayout(add_selected_layout)

        # Cart section
        layout.addWidget(QLabel("<b>Order Cart</b>"))
        self.cart_table = QTableWidget(0, 5)
        self.cart_table.setHorizontalHeaderLabels(
            ["SKU", "Brand", "Product", "Qty", "Remove"]
        )
        header = self.cart_table.horizontalHeader()
        for col in range(self.cart_table.columnCount()):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        self.cart_table.setColumnWidth(0, 140)
        self.cart_table.setColumnWidth(1, 120)
        self.cart_table.setColumnWidth(2, 420)
        self.cart_table.setColumnWidth(3, 80)
        layout.addWidget(self.cart_table)
        
        # Total and buttons
        bottom_layout = QHBoxLayout()
        self.total_label = QLabel("Cart Total: EUR 0.00")
        self.total_label.setStyleSheet("font-weight: bold; font-size: 12pt;")
        bottom_layout.addWidget(self.total_label)
        bottom_layout.addStretch()
        
        self.notes_input = QTextEdit()
        self.notes_input.setPlaceholderText("Order notes or instructions")
        self.notes_input.setMaximumHeight(80)
        
        button_layout = QVBoxLayout()
        button_layout.addWidget(QLabel("Notes:"))
        button_layout.addWidget(self.notes_input)
        
        button_row = QHBoxLayout()
        self.save_button = QPushButton("Save Order")
        self.save_button.clicked.connect(self.on_save_order)
        button_row.addWidget(self.save_button)
        
        self.pdf_button = QPushButton("Export to PDF")
        self.pdf_button.clicked.connect(self.on_export_to_pdf)
        self.pdf_button.setEnabled(False)  # disabled until order is saved
        button_row.addWidget(self.pdf_button)
        
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_row.addWidget(self.cancel_button)
        
        button_layout.addLayout(button_row)
        bottom_layout.addLayout(button_layout)
        layout.addLayout(bottom_layout)

    def _populate_brand_filter(self) -> None:
        brands = set()
        for product in self.products:
            brand = product.get("brand")
            if brand:
                brands.add(str(brand))
        for brand in sorted(brands):
            self.brand_filter.addItem(brand)

    def _load_products_table(self) -> None:
        self.products_table.setRowCount(len(self.products))
        self.product_row_map = {}  # Map row index to product
        for row, product in enumerate(self.products):
            self.product_row_map[row] = product
            
            # SKU
            sku = str(product.get("sku") or "")
            self.products_table.setItem(row, 0, QTableWidgetItem(sku))
            
            # Brand
            brand = str(product.get("brand") or "")
            self.products_table.setItem(row, 1, QTableWidgetItem(brand))
            
            # Product name
            product_name = str(product.get("product_name") or "")
            self.products_table.setItem(row, 2, QTableWidgetItem(product_name))
            
            # Pack
            pack = str(product.get("pack") or "")
            self.products_table.setItem(row, 3, QTableWidgetItem(pack))
            
            # Unit
            unit = str(product.get("unit") or "")
            self.products_table.setItem(row, 4, QTableWidgetItem(unit))
            
            # CTN Qty
            ctn_qty = str(product.get("ctn_qty") or 0)
            self.products_table.setItem(row, 5, QTableWidgetItem(ctn_qty))
            
            # Price
            price = float(product.get("price") or 0.0)
            self.products_table.setItem(row, 6, QTableWidgetItem(f"{price:.2f}"))
            
            # Qty dropdown (0-10)
            qty_combo = QComboBox()
            qty_combo.addItems([str(i) for i in range(11)])
            qty_combo.setCurrentText("0")
            self.products_table.setCellWidget(row, 7, qty_combo)
            
            # Custom qty input
            custom_qty = QSpinBox()
            custom_qty.setMinimum(0)
            custom_qty.setMaximum(10000)
            custom_qty.setValue(0)
            self.products_table.setCellWidget(row, 8, custom_qty)
            
            # Add to cart button
            add_btn = QPushButton("Add")
            add_btn.clicked.connect(lambda checked, r=row: self.on_add_to_cart(r))
            self.products_table.setCellWidget(row, 9, add_btn)

    def _filter_products(self) -> None:
        search_text = self.search_input.text().lower()
        selected_brand = self.brand_filter.currentText()
        
        for row in range(self.products_table.rowCount()):
            product = self.product_row_map.get(row)
            if not product:
                continue
            
            # Check search text match
            product_name = str(product.get("product_name") or "").lower()
            matches_search = search_text == "" or search_text in product_name
            
            # Check brand filter match
            brand = str(product.get("brand") or "")
            matches_brand = selected_brand == "All brands" or brand == selected_brand
            
            # Show/hide row
            show_row = matches_search and matches_brand
            self.products_table.setRowHidden(row, not show_row)

    def _get_row_quantity(self, row: int) -> int:
        qty_combo = self.products_table.cellWidget(row, 7)
        custom_qty = self.products_table.cellWidget(row, 8)
        qty_0_10 = int(qty_combo.currentText() or 0) if qty_combo else 0
        custom_qty_val = custom_qty.value() if custom_qty else 0
        return custom_qty_val if custom_qty_val > 10 else qty_0_10

    def _set_row_quantity_to_zero(self, row: int) -> None:
        qty_combo = self.products_table.cellWidget(row, 7)
        custom_qty = self.products_table.cellWidget(row, 8)
        if qty_combo:
            qty_combo.setCurrentText("0")
        if custom_qty:
            custom_qty.setValue(0)

    def _cart_item_key(self, product: Dict[str, any]) -> tuple:
        sku = str(product.get("sku") or "").strip()
        if sku:
            return ("sku", sku)
        return (
            "name",
            str(product.get("product_name") or "").strip().lower(),
            str(product.get("pack") or "").strip().lower(),
            str(product.get("unit") or "").strip().lower(),
        )

    def _find_duplicate_cart_item(self, product: Dict[str, any]) -> Optional[Dict[str, any]]:
        key = self._cart_item_key(product)
        for item in self.cart_items:
            if self._cart_item_key(item) == key:
                return item
        return None

    def _has_duplicate_cart_items(self) -> Optional[str]:
        seen = set()
        for item in self.cart_items:
            key = self._cart_item_key(item)
            if key in seen:
                return str(item.get("sku") or item.get("product_name") or "product")
            seen.add(key)
        return None

    def _create_cart_item(self, product: Dict[str, any], quantity: int) -> Dict[str, any]:
        return {
            "sku": product.get("sku"),
            "product_name": product.get("product_name"),
            "brand": product.get("brand"),
            "pack": product.get("pack"),
            "unit": product.get("unit"),
            "ctn_qty": product.get("ctn_qty") or 0,
            "unit_price": product.get("price") or 0.0,
            "quantity": quantity,
            "total_price": quantity * (product.get("ctn_qty") or 1) * (product.get("price") or 0.0),
            "extra_json": product.get("extra_json"),
        }

    def on_add_to_cart(self, row: int) -> None:
        product = self.product_row_map.get(row)
        if not product:
            return
        quantity = self._get_row_quantity(row)
        if quantity <= 0:
            QMessageBox.warning(self, "Invalid quantity", "Please select a quantity > 0")
            return
        if self._find_duplicate_cart_item(product):
            QMessageBox.warning(
                self,
                "Duplicate item",
                "This product is already in the cart. Please update the existing line instead of adding it again.",
            )
            return
        cart_item = self._create_cart_item(product, quantity)
        self.cart_items.append(cart_item)
        self._update_cart_table()
        self._set_row_quantity_to_zero(row)
        QMessageBox.information(self, "Added", f"Added {quantity} unit(s) to cart")

    def on_add_selected_row_to_cart(self) -> None:
        row = self.selected_product_row
        if row is None or row < 0 or row >= self.products_table.rowCount():
            QMessageBox.warning(self, "No row selected", "Please select a row to add to cart.")
            return
        product = self.product_row_map.get(row)
        if not product:
            QMessageBox.warning(self, "Invalid row", "Selected row does not contain a valid product.")
            return
        quantity = self._get_row_quantity(row)
        if quantity <= 0:
            QMessageBox.warning(self, "Invalid quantity", "Please choose a valid quantity for the selected row.")
            return
        if self._find_duplicate_cart_item(product):
            QMessageBox.warning(
                self,
                "Duplicate item",
                "This product is already in the cart. Please update the existing line instead of adding it again.",
            )
            return
        cart_item = self._create_cart_item(product, quantity)
        self.cart_items.append(cart_item)
        self._update_cart_table()
        self._set_row_quantity_to_zero(row)
        QMessageBox.information(self, "Added", f"Added {quantity} unit(s) to cart")

    def _update_selected_row(self) -> None:
        self.selected_product_row = self.products_table.currentRow()

    def _update_cart_table(self) -> None:
        self.cart_table.setRowCount(len(self.cart_items))
        total_amount = 0.0
        
        for row, item in enumerate(self.cart_items):
            # SKU
            self.cart_table.setItem(row, 0, QTableWidgetItem(str(item.get("sku") or "")))
            
            # Brand
            self.cart_table.setItem(row, 1, QTableWidgetItem(str(item.get("brand") or "")))
            
            # Product
            self.cart_table.setItem(row, 2, QTableWidgetItem(str(item.get("product_name") or "")))
            
            # Qty
            self.cart_table.setItem(row, 3, QTableWidgetItem(str(item.get("quantity") or 0)))
            
            total = float(item.get("total_price") or 0.0)
            total_amount += total
            
            # Remove button
            remove_btn = QPushButton("Remove")
            remove_btn.clicked.connect(lambda checked, r=row: self.on_remove_from_cart(r))
            self.cart_table.setCellWidget(row, 4, remove_btn)
        
        self.total_label.setText(f"Cart Total: EUR {total_amount:.2f}")

    def on_remove_from_cart(self, row: int) -> None:
        if 0 <= row < len(self.cart_items):
            self.cart_items.pop(row)
            self._update_cart_table()

    def on_save_order(self) -> None:
        if not self.cart_items:
            QMessageBox.warning(self, "Empty cart", "Please add items to cart before saving order.")
            return
        self._save_order_to_excel_and_db()

    def on_export_to_pdf(self) -> None:
        if not self.cart_items:
            QMessageBox.warning(self, "Empty cart", "Please add items to cart before exporting.")
            return
        self._export_to_pdf()

    def on_generate_order(self, format: str = "excel") -> None:
        if not self.cart_items:
            QMessageBox.warning(self, "Empty cart", "Please add items to cart before generating order.")
            return
        
        if format == "excel":
            self._save_order_to_excel_and_db()
        elif format == "pdf":
            self._export_to_pdf()
    
    def _save_order_to_excel_and_db(self) -> None:
        filename = f"{self.vendor}_order_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        order_dir = VENDOR_ROOT / self.vendor / "orders"
        order_dir.mkdir(parents=True, exist_ok=True)
        order_path = order_dir / filename
        
        try:
            from pandas import DataFrame

            order_data = []
            for item in self.cart_items:
                order_data.append({
                    "Vendor": self.vendor,
                    "SKU": item.get("sku"),
                    "Product": item.get("product_name"),
                    "Brand": item.get("brand"),
                    "Quantity": item.get("quantity"),
                })

            DataFrame(order_data).to_excel(order_path, index=False)
        except Exception as exc:
            QMessageBox.critical(self, "Order creation failed", f"Could not create order Excel sheet:\n{exc}")
            return

        # Prevent duplicate cart entries before saving
        duplicate_item = self._has_duplicate_cart_items()
        if duplicate_item:
            QMessageBox.warning(
                self,
                "Duplicate items",
                f"Duplicate product '{duplicate_item}' found in the cart. Please remove or merge duplicates before saving.",
            )
            return

        # Save or update database
        try:
            total_amount = sum(item.get("total_price", 0.0) for item in self.cart_items)
            if getattr(self, "order_id", None):
                # update existing order
                try:
                    self.db.update_order(
                        self.order_id,
                        self.cart_items,
                        total_amount=total_amount,
                        notes=self.notes_input.toPlainText().strip(),
                        order_filename=str(order_path),
                    )
                except Exception:
                    # fallback to save as new order if update fails
                    self.db.save_order(
                        vendor=self.vendor,
                        items=self.cart_items,
                        total_amount=total_amount,
                        order_filename=str(order_path),
                        notes=self.notes_input.toPlainText().strip(),
                    )
            else:
                self.order_id = self.db.save_order(
                    vendor=self.vendor,
                    items=self.cart_items,
                    total_amount=total_amount,
                    order_filename=str(order_path),
                    notes=self.notes_input.toPlainText().strip(),
                )
        except Exception as exc:
            QMessageBox.warning(self, "Database save", f"Order Excel saved but DB save failed:\n{exc}")

        QMessageBox.information(self, "Order saved", f"Order saved successfully. You can now export to PDF.")
        self.pdf_button.setEnabled(True)
        if hasattr(self.parent(), "_load_order_history"):
            try:
                self.parent()._load_order_history()
            except Exception:
                pass

    def _export_to_excel(self) -> None:
        filename = f"{self.vendor}_order_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        order_dir = VENDOR_ROOT / self.vendor / "orders"
        order_dir.mkdir(parents=True, exist_ok=True)
        order_path = order_dir / filename

        try:
            from pandas import DataFrame

            order_data = []
            for item in self.cart_items:
                order_data.append({
                    "Vendor": self.vendor,
                    "SKU": item.get("sku"),
                    "Product": item.get("product_name"),
                    "Brand": item.get("brand"),
                    "Quantity": item.get("quantity"),
                })

            DataFrame(order_data).to_excel(order_path, index=False)
        except Exception as exc:
            QMessageBox.critical(self, "Excel export failed", f"Could not create Excel file:\n{exc}")
            return

        QMessageBox.information(self, "Excel exported", f"Order Excel created:\n{order_path}")

    def _export_to_pdf(self) -> None:
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib import colors
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch
        except ImportError:
            QMessageBox.warning(self, "Missing dependency", "Please install reportlab: pip install reportlab")
            return
        
        filename = f"{self.vendor}_order_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        order_dir = VENDOR_ROOT / self.vendor / "orders"
        order_dir.mkdir(parents=True, exist_ok=True)
        pdf_path = order_dir / filename
        
        try:
            doc = SimpleDocTemplate(str(pdf_path), pagesize=A4)
            elements = []
            styles = getSampleStyleSheet()

            # Header with logo and company details
            header_style = ParagraphStyle('Header', parent=styles['Normal'], fontSize=10)
            title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=16, leading=18)
            # build header table
            header_data = []
            logo_path = LOGO_FILE if LOGO_FILE and LOGO_FILE.exists() else None
            left = []
            if logo_path:
                from reportlab.platypus import Image

                try:
                    img = Image(str(logo_path), width=80, height=40)
                    left.append(img)
                except Exception:
                    left.append(Paragraph(COMPANY_NAME, header_style))
            else:
                left.append(Paragraph(COMPANY_NAME, header_style))

            right_lines = [COMPANY_NAME, COMPANY_ADDRESS, f"Tax ID: 01435901405", f"VAT ID: DE365100311"]
            right = Paragraph("<br/>".join(right_lines), header_style)
            header_data.append([left, right])
            header_table = Table(header_data, colWidths=[2.0 * inch, 4.5 * inch])
            header_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
            elements.append(header_table)
            elements.append(Spacer(1, 0.15 * inch))

            profile = self.db.get_vendor_profile(self.vendor)
            if profile:
                profile_lines = [f"<b>Vendor profile</b>"]
                if profile.get("display_name"):
                    profile_lines.append(f"Name: {profile.get('display_name')}")
                if profile.get("customer_number"):
                    profile_lines.append(f"Customer #: {profile.get('customer_number')}")
                
                
                if profile.get("customer_website"):
                    profile_lines.append(f"Website: {profile.get('customer_website')}")
                if profile.get("country_of_origin"):
                    profile_lines.append(f"Country of origin: {profile.get('country_of_origin')}")
                if profile.get("address"):
                    profile_lines.append(f"Address: {profile.get('address')}")
                if profile.get("iban"):
                    profile_lines.append(f"IBAN: {profile.get('iban')}")
                if profile.get("bank_name"):
                    profile_lines.append(f"Bank: {profile.get('bank_name')}")
                if profile.get("swift_bic"):
                    profile_lines.append(f"SWIFT/BIC: {profile.get('swift_bic')}")
                elements.append(Paragraph("<br/>".join(profile_lines), styles['Normal']))
                elements.append(Spacer(1, 0.2 * inch))

            elements.append(Paragraph(f"Order For {self.vendor}", title_style))
            elements.append(Paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
            elements.append(Spacer(1, 0.2 * inch))

            # Table data (no Unit column per request)
            data = [["SKU", "Brand", "Product", "Qty"]]
            for item in self.cart_items:
                data.append([
                    str(item.get("sku") or ""),
                    str(item.get("brand") or ""),
                    str(item.get("product_name") or ""),
                    str(item.get("quantity") or 0),
                ])

            table = Table(data, colWidths=[1.2 * inch, 1.2 * inch, 3.6 * inch, 0.8 * inch])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            ]))
            elements.append(table)

            # Total
            elements.append(Spacer(1, 0.2 * inch))
            total_amount = sum(item.get("total_price", 0.0) for item in self.cart_items)
            elements.append(Paragraph(f"<b>Total: EUR {total_amount:.2f}</b>", styles['Normal']))

            # Notes
            notes = self.notes_input.toPlainText().strip()
            if notes:
                elements.append(Spacer(1, 0.2 * inch))
                elements.append(Paragraph("<b>Notes:</b>", styles['Normal']))
                elements.append(Paragraph(notes, styles['Normal']))

            doc.build(elements)
        except Exception as exc:
            QMessageBox.critical(self, "PDF export failed", f"Could not create PDF:\n{exc}")
            return
        
        QMessageBox.information(self, "PDF exported", f"Order PDF created:\n{pdf_path}")
        self.accept()


class VendorProfileDialog(QDialog):
    def __init__(self, vendor: str, profile: Dict[str, Any], db: DatabaseManager, parent=None):
        super().__init__(parent)
        self.vendor = vendor
        self.profile = profile or {}
        self.db = db
        self.setWindowTitle(f"Vendor details - {vendor}")
        self.setMinimumSize(560, 520)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.display_name_input = QLineEdit(str(self.profile.get("display_name") or ""))
        self.legal_name_input = QLineEdit(str(self.profile.get("legal_name") or ""))
        self.country_of_origin_input = QLineEdit(str(self.profile.get("country_of_origin") or ""))
        self.address_input = QTextEdit(str(self.profile.get("address") or ""))
        self.address_input.setMaximumHeight(100)
        self.contact_name_input = QLineEdit(str(self.profile.get("contact_name") or ""))
        self.contact_email_input = QLineEdit(str(self.profile.get("contact_email") or ""))
        self.contact_phone_input = QLineEdit(str(self.profile.get("contact_phone") or ""))
        self.tax_id_input = QLineEdit(str(self.profile.get("tax_id") or ""))
        self.vat_id_input = QLineEdit(str(self.profile.get("vat_id") or ""))
        self.iban_input = QLineEdit(str(self.profile.get("iban") or ""))
        self.bank_name_input = QLineEdit(str(self.profile.get("bank_name") or ""))
        self.swift_bic_input = QLineEdit(str(self.profile.get("swift_bic") or ""))
        self.customer_number_input = QLineEdit(str(self.profile.get("customer_number") or ""))
        self.customer_website_input = QLineEdit(str(self.profile.get("customer_website") or ""))

        self.display_name_input.setPlaceholderText("Optional display name for reports")
        self.country_of_origin_input.setPlaceholderText("Country of origin / supplier country")
        self.iban_input.setPlaceholderText("International Bank Account Number")
        self.swift_bic_input.setPlaceholderText("SWIFT / BIC code")
        self.customer_number_input.setPlaceholderText("Customer number for this vendor")
        self.customer_website_input.setPlaceholderText("Vendor website (https://...)")

        form.addRow("Vendor folder", QLabel(self.vendor))
        form.addRow("Display name", self.display_name_input)
        form.addRow("Legal name", self.legal_name_input)
        form.addRow("Country of origin", self.country_of_origin_input)
        form.addRow("Customer number", self.customer_number_input)
        form.addRow("Website", self.customer_website_input)
        form.addRow("Address", self.address_input)
        form.addRow("Contact name", self.contact_name_input)
        form.addRow("Contact email", self.contact_email_input)
        form.addRow("Contact phone", self.contact_phone_input)
        form.addRow("Tax ID", self.tax_id_input)
        form.addRow("VAT ID", self.vat_id_input)
        form.addRow("IBAN", self.iban_input)
        form.addRow("Bank name", self.bank_name_input)
        form.addRow("SWIFT/BIC", self.swift_bic_input)

        layout.addLayout(form)

        button_row = QHBoxLayout()
        save_btn = QPushButton("Save details")
        save_btn.clicked.connect(self.on_save)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        button_row.addStretch()
        button_row.addWidget(save_btn)
        button_row.addWidget(cancel_btn)
        layout.addLayout(button_row)

    def on_save(self) -> None:
        profile = {
            "display_name": self.display_name_input.text().strip(),
            "legal_name": self.legal_name_input.text().strip(),
            "country_of_origin": self.country_of_origin_input.text().strip(),
            "address": self.address_input.toPlainText().strip(),
            "contact_name": self.contact_name_input.text().strip(),
            "contact_email": self.contact_email_input.text().strip(),
            "contact_phone": self.contact_phone_input.text().strip(),
            "tax_id": self.tax_id_input.text().strip(),
            "vat_id": self.vat_id_input.text().strip(),
            "iban": self.iban_input.text().strip(),
            "bank_name": self.bank_name_input.text().strip(),
            "swift_bic": self.swift_bic_input.text().strip(),
            "customer_number": self.customer_number_input.text().strip(),
                "customer_website": self.customer_website_input.text().strip(),
        }
        try:
            self.db.save_vendor_profile(self.vendor, profile)
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", f"Could not save vendor details:\n{exc}")
            return
        QMessageBox.information(self, "Saved", "Vendor details saved successfully.")
        self.accept()

class EditProductDialog(QDialog):
    def __init__(self, product: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.product = product
        self.setWindowTitle(f"Edit Product: {product.get('sku')}")
        self.setMinimumSize(400, 350)
        self.updated_data = {}
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.sku_input = QLineEdit(str(self.product.get("sku") or ""))
        self.sku_input.setReadOnly(True)
        self.product_name_input = QLineEdit(str(self.product.get("product_name") or ""))
        self.brand_input = QLineEdit(str(self.product.get("brand") or ""))
        self.pack_input = QLineEdit(str(self.product.get("pack") or ""))
        self.unit_input = QLineEdit(str(self.product.get("unit") or ""))
        
        self.ctn_qty_input = QSpinBox()
        self.ctn_qty_input.setMinimum(0)
        self.ctn_qty_input.setMaximum(100000)
        self.ctn_qty_input.setValue(int(self.product.get("ctn_qty") or 0))

        self.price_input = QLineEdit(str(self.product.get("price") or 0.0))
        self.stock_input = QLineEdit(str(self.product.get("stock") or ""))
        
        self.barcode_input = QLineEdit(str(self.product.get("barcode") or ""))
        self.sr_sku_input = QLineEdit(str(self.product.get("sr_sku") or ""))

        form.addRow("SKU", self.sku_input)
        form.addRow("Product Name", self.product_name_input)
        form.addRow("Brand", self.brand_input)
        form.addRow("Pack", self.pack_input)
        form.addRow("Unit", self.unit_input)
        form.addRow("CTN Qty", self.ctn_qty_input)
        form.addRow("Price", self.price_input)
        form.addRow("Stock (Status/Qty)", self.stock_input)
        form.addRow("Barcode", self.barcode_input)
        form.addRow("SR-SKU", self.sr_sku_input)

        layout.addLayout(form)

        button_row = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.on_save)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        button_row.addStretch()
        button_row.addWidget(save_btn)
        button_row.addWidget(cancel_btn)
        layout.addLayout(button_row)

    def on_save(self) -> None:
        try:
            price_val = float(self.price_input.text())
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Price must be a valid number.")
            return

        self.updated_data = {
            "product_name": self.product_name_input.text().strip(),
            "brand": self.brand_input.text().strip(),
            "pack": self.pack_input.text().strip(),
            "unit": self.unit_input.text().strip(),
            "ctn_qty": self.ctn_qty_input.value(),
            "price": price_val,
            "stock": self.stock_input.text().strip(),
            "barcode": self.barcode_input.text().strip(),
            "sr_sku": self.sr_sku_input.text().strip(),
            "last_updated": datetime.now().isoformat()
        }
        self.accept()


class InventoryDialog(QDialog):
    def __init__(self, current_vendor: str, all_vendors: List[str], db: DatabaseManager, parent=None):
        super().__init__(parent)
        self.current_vendor = current_vendor
        self.all_vendors = all_vendors
        self.db = db
        self.products: List[Dict[str, Any]] = []
        self.setWindowTitle("Product Inventory Management")
        self.setMinimumSize(1200, 800)
        self._build_ui()
        self._load_products()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)

        # Header controls
        header_layout = QHBoxLayout()
        
        header_layout.addWidget(QLabel("Vendor:"))
        self.vendor_combo = QComboBox()
        self.vendor_combo.addItems(self.all_vendors)
        self.vendor_combo.setCurrentText(self.current_vendor)
        self.vendor_combo.currentTextChanged.connect(self.on_vendor_changed)
        header_layout.addWidget(self.vendor_combo)

        header_layout.addWidget(QLabel("Search:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search by name or SKU...")
        self.search_input.textChanged.connect(self._filter_products)
        header_layout.addWidget(self.search_input)

        header_layout.addWidget(QLabel("Brand:"))
        self.brand_filter = QComboBox()
        self.brand_filter.addItem("All brands")
        self.brand_filter.currentTextChanged.connect(self._filter_products)
        header_layout.addWidget(self.brand_filter)

        header_layout.addStretch()

        self.export_vendor_btn = QPushButton("Export Current Vendor")
        self.export_vendor_btn.clicked.connect(self.on_export_vendor)
        header_layout.addWidget(self.export_vendor_btn)

        self.export_all_btn = QPushButton("Export All Vendors")
        self.export_all_btn.clicked.connect(self.on_export_all)
        header_layout.addWidget(self.export_all_btn)

        layout.addLayout(header_layout)

        # Table
        self.table = QTableWidget(0, 10)
        self.table.setHorizontalHeaderLabels(
            ["SKU", "Product", "Brand", "Pack", "Unit", "CTN Qty", "Price", "Stock", "Barcode", "SR-SKU"]
        )
        self.table.setSortingEnabled(True)
        header = self.table.horizontalHeader()
        for col in range(self.table.columnCount()):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setColumnWidth(1, 400)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.itemDoubleClicked.connect(self.on_edit_selected)
        layout.addWidget(self.table)

        # Bottom controls
        bottom_layout = QHBoxLayout()
        self.edit_btn = QPushButton("Edit Selected Product")
        self.edit_btn.clicked.connect(self.on_edit_selected)
        bottom_layout.addWidget(self.edit_btn)
        bottom_layout.addStretch()
        self.close_btn = QPushButton("Close")
        self.close_btn.clicked.connect(self.accept)
        bottom_layout.addWidget(self.close_btn)

        layout.addLayout(bottom_layout)

    def on_vendor_changed(self, vendor_name: str) -> None:
        self.current_vendor = vendor_name
        self._load_products()

    def _load_products(self) -> None:
        self.table.setSortingEnabled(False)
        self.products = self.db.get_vendor_products(self.current_vendor)
        
        self.brand_filter.blockSignals(True)
        self.brand_filter.clear()
        self.brand_filter.addItem("All brands")
        brands = sorted({str(p.get("brand")).strip() for p in self.products if p.get("brand")})
        for b in brands:
            if b:
                self.brand_filter.addItem(b)
        self.brand_filter.blockSignals(False)
        
        self.table.setRowCount(len(self.products))
        
        for row, product in enumerate(self.products):
            sku_item = QTableWidgetItem(str(product.get("sku") or ""))
            sku_item.setData(Qt.UserRole, product)
            
            self.table.setItem(row, 0, sku_item)
            self.table.setItem(row, 1, QTableWidgetItem(str(product.get("product_name") or "")))
            self.table.setItem(row, 2, QTableWidgetItem(str(product.get("brand") or "")))
            self.table.setItem(row, 3, QTableWidgetItem(str(product.get("pack") or "")))
            self.table.setItem(row, 4, QTableWidgetItem(str(product.get("unit") or "")))
            
            ctn_qty_item = QTableWidgetItem()
            ctn_qty_item.setData(Qt.DisplayRole, int(product.get("ctn_qty") or 0))
            self.table.setItem(row, 5, ctn_qty_item)
            
            price_item = QTableWidgetItem()
            price_item.setData(Qt.DisplayRole, float(product.get("price") or 0.0))
            self.table.setItem(row, 6, price_item)
            
            self.table.setItem(row, 7, QTableWidgetItem(str(product.get("stock") or "")))
            self.table.setItem(row, 8, QTableWidgetItem(str(product.get("barcode") or "")))
            self.table.setItem(row, 9, QTableWidgetItem(str(product.get("sr_sku") or "")))
            
        self.table.setSortingEnabled(True)
        self._filter_products()

    def _filter_products(self) -> None:
        search_text = self.search_input.text().lower()
        selected_brand = self.brand_filter.currentText()
        
        for row in range(self.table.rowCount()):
            sku_item = self.table.item(row, 0)
            if not sku_item:
                continue
            
            product = sku_item.data(Qt.UserRole)
            if not product:
                continue
            
            prod_name = str(product.get("product_name") or "").lower()
            sku = str(product.get("sku") or "").lower()
            matches_search = search_text == "" or search_text in prod_name or search_text in sku
            
            brand = str(product.get("brand") or "").strip()
            matches_brand = selected_brand == "All brands" or brand == selected_brand
            
            self.table.setRowHidden(row, not (matches_search and matches_brand))

    def on_edit_selected(self) -> None:
        sel_items = self.table.selectedItems()
        if not sel_items:
            QMessageBox.warning(self, "No selection", "Please select a product to edit.")
            return
        
        row = sel_items[0].row()
        sku_item = self.table.item(row, 0)
        product = sku_item.data(Qt.UserRole)
        
        dialog = EditProductDialog(product, self)
        if dialog.exec_():
            updated_data = dialog.updated_data
            sku = product.get("sku")
            if sku:
                try:
                    self.db.update_vendor_product(self.current_vendor, sku, updated_data)
                    self._load_products()
                except Exception as exc:
                    QMessageBox.critical(self, "Error", f"Failed to update product:\n{exc}")

    def on_export_vendor(self) -> None:
        if not self.products:
            QMessageBox.warning(self, "No data", "No products to export.")
            return
        self._export_to_excel(self.products, f"{self.current_vendor}_inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

    def on_export_all(self) -> None:
        all_products = []
        for vendor in self.all_vendors:
            prods = self.db.get_vendor_products(vendor)
            for p in prods:
                p["_vendor"] = vendor
            all_products.extend(prods)
        
        if not all_products:
            QMessageBox.warning(self, "No data", "No products to export.")
            return
        self._export_to_excel(all_products, f"all_vendors_inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

    def _export_to_excel(self, data_list: List[Dict[str, Any]], filename: str) -> None:
        try:
            from pandas import DataFrame
            
            export_data = []
            for item in data_list:
                row_dict = {
                    "Vendor": item.get("_vendor", self.current_vendor),
                    "SKU": item.get("sku"),
                    "Product": item.get("product_name"),
                    "Brand": item.get("brand"),
                    "Pack": item.get("pack"),
                    "Unit": item.get("unit"),
                    "CTN Qty": item.get("ctn_qty"),
                    "Price": item.get("price"),
                    "Stock": item.get("stock"),
                    "Barcode": item.get("barcode"),
                    "SR-SKU": item.get("sr_sku"),
                }
                export_data.append(row_dict)
                
            if "all_vendors" in filename:
                export_dir = VENDOR_ROOT / "exports"
            else:
                export_dir = VENDOR_ROOT / self.current_vendor / "exports"
            export_dir.mkdir(parents=True, exist_ok=True)
            export_path = export_dir / filename
            
            DataFrame(export_data).to_excel(export_path, index=False)
            QMessageBox.information(self, "Export Successful", f"Inventory exported to:\n{export_path}")
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", f"Could not create Excel file:\n{exc}")

class VendorManagerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sunrise Supermarket Vendor Manager")
        self.setMinimumSize(1200, 760)
        self.db = DatabaseManager(DB_FILE)
        self.loader = ExcelLoader(VENDOR_ROOT)
        self.current_vendor: Optional[str] = None
        self.vendor_names: List[str] = []
        self.product_cache: List[Dict] = []
        self._build_ui()
        self._load_sidebar()
        self._refresh_all_stats()

    def _build_ui(self) -> None:
        root = QWidget()
        main_layout = QHBoxLayout(root)
        self.setCentralWidget(root)

        left_panel = QVBoxLayout()
        self.vendor_list = QListWidget()
        self.vendor_list.itemSelectionChanged.connect(self.on_vendor_selected)
        left_panel.addWidget(QLabel("Vendor list"))
        left_panel.addWidget(self.vendor_list)
        self.sync_vendor_button = QPushButton("Sync selected vendor")
        self.sync_vendor_button.clicked.connect(self.sync_selected_vendor)
        self.sync_all_button = QPushButton("Sync all vendors")
        self.sync_all_button.clicked.connect(self.sync_all_vendors)
        self.reset_vendor_button = QPushButton("Reset selected vendor products")
        self.reset_vendor_button.clicked.connect(self.reset_selected_vendor_products)
        self.inventory_button = QPushButton("Product inventory")
        self.inventory_button.clicked.connect(self.open_inventory_dialog)
        self.order_button = QPushButton("Create order")
        self.order_button.clicked.connect(self.open_order_dialog)
        self.vendor_details_button = QPushButton("Vendor details")
        self.vendor_details_button.clicked.connect(self.open_vendor_details)
        left_panel.addWidget(self.sync_vendor_button)
        left_panel.addWidget(self.sync_all_button)
        left_panel.addWidget(self.reset_vendor_button)
        left_panel.addWidget(self.inventory_button)
        left_panel.addWidget(self.order_button)
        left_panel.addWidget(self.vendor_details_button)

        right_panel = QVBoxLayout()
        header = QHBoxLayout()
        logo_label = QLabel()
        logo_label.setFixedSize(120, 100)
        if LOGO_FILE.exists():
            logo_label.setPixmap(QPixmap(str(LOGO_FILE)).scaled(120, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        header.addWidget(logo_label)
        company_box = QVBoxLayout()
        company_box.addWidget(QLabel(f"<h1>{COMPANY_NAME}</h1>"))
        company_box.addWidget(QLabel(COMPANY_ADDRESS))
        company_box.addStretch()
        header.addLayout(company_box)
        header.addStretch()
        right_panel.addLayout(header)

        stats_box = QGroupBox("Statistics")
        stats_layout = QVBoxLayout()
        self.stats_label = QLabel("Select a vendor or sync products to see statistics.")
        self.stats_table = QTableWidget(0, 5)
        self.stats_table.setHorizontalHeaderLabels(["Vendor", "Products", "Inventory value", "Barcodes", "SR-SKUs"])
        self.stats_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        stats_layout.addWidget(self.stats_label)
        stats_layout.addWidget(self.stats_table)
        stats_box.setLayout(stats_layout)
        right_panel.addWidget(stats_box)

        details_box = QGroupBox("Vendor business details")
        details_layout = QVBoxLayout()
        self.vendor_detail_summary = QLabel("No vendor selected.")
        self.vendor_detail_summary.setWordWrap(True)
        self.vendor_detail_summary.setTextFormat(Qt.RichText)
        details_layout.addWidget(self.vendor_detail_summary)
        details_box.setLayout(details_layout)
        right_panel.addWidget(details_box)

        self.product_table = QTableWidget(0, 10)
        self.product_table.setHorizontalHeaderLabels(
            ["SKU", "Product", "Brand", "Pack", "Unit", "CTN Qty", "Price", "Stock", "Barcode", "SR-SKU"]
        )
        header = self.product_table.horizontalHeader()
        for col in range(self.product_table.columnCount()):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        self.product_table.setColumnWidth(1, 520)
        self.product_table.setColumnWidth(3, 120)
        self.product_table.setColumnWidth(4, 80)
        self.product_table.setColumnWidth(5, 90)
        self.product_table.setColumnWidth(8, 140)
        self.product_table.setColumnWidth(9, 180)
        right_panel.addWidget(self.product_table)

        orders_box = QGroupBox("Order history")
        orders_layout = QVBoxLayout()
        # action buttons for selected order
        orders_btn_row = QHBoxLayout()
        self.open_order_btn = QPushButton("Open/Edit Selected")
        self.open_order_btn.clicked.connect(self.on_open_selected_order)
        orders_btn_row.addWidget(self.open_order_btn)
        self.export_order_excel_btn = QPushButton("Export Selected (Excel)")
        self.export_order_excel_btn.clicked.connect(lambda: self.on_export_selected_order("excel"))
        orders_btn_row.addWidget(self.export_order_excel_btn)
        self.export_order_pdf_btn = QPushButton("Export Selected (PDF)")
        self.export_order_pdf_btn.clicked.connect(lambda: self.on_export_selected_order("pdf"))
        orders_btn_row.addWidget(self.export_order_pdf_btn)
        self.delete_order_btn = QPushButton("Delete Selected")
        self.delete_order_btn.clicked.connect(self.on_delete_selected_order)
        orders_btn_row.addWidget(self.delete_order_btn)
        # demo order button (for testing)
        self.create_demo_btn = QPushButton("Create Demo Order")
        self.create_demo_btn.clicked.connect(self.on_create_demo_order)
        orders_btn_row.addWidget(self.create_demo_btn)
        orders_layout.addLayout(orders_btn_row)

        self.orders_table = QTableWidget(0, 6)
        self.orders_table.setHorizontalHeaderLabels(["Vendor", "Created", "Updated", "Status", "Total", "Filename"])
        self.orders_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.orders_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.orders_table.setSelectionMode(QTableWidget.SingleSelection)
        self.orders_table.itemSelectionChanged.connect(self.on_order_selection_changed)
        self.orders_table.cellDoubleClicked.connect(lambda r, c: self.on_open_selected_order())
        orders_layout.addWidget(self.orders_table)
        orders_box.setLayout(orders_layout)
        right_panel.addWidget(orders_box)

        main_layout.addLayout(left_panel, 2)
        main_layout.addLayout(right_panel, 6)

    def _load_sidebar(self) -> None:
        self.vendor_list.clear()
        self.vendor_names = [p.name for p in VENDOR_ROOT.iterdir() if p.is_dir()]
        self.vendor_names.sort()
        for vendor in self.vendor_names:
            self.vendor_list.addItem(vendor)

    def _refresh_all_stats(self) -> None:
        stats = self.db.get_all_vendor_statistics(self.vendor_names)
        self.stats_table.setRowCount(0)
        total_products = 0
        total_value = 0.0
        for row in stats:
            index = self.stats_table.rowCount()
            self.stats_table.insertRow(index)
            self.stats_table.setItem(index, 0, QTableWidgetItem(row["vendor"]))
            self.stats_table.setItem(index, 1, QTableWidgetItem(str(row["product_count"])))
            self.stats_table.setItem(index, 2, QTableWidgetItem(f"{row['inventory_value']:.2f}"))
            self.stats_table.setItem(index, 3, QTableWidgetItem(str(row.get("barcode_count") or 0)))
            self.stats_table.setItem(index, 4, QTableWidgetItem(str(row.get("srsku_count") or 0)))
            total_products += row["product_count"]
            total_value += row["inventory_value"]
        self.stats_label.setText(f"{len(stats)} vendors, {total_products} total products, inventory value €{total_value:.2f}")
        self._load_order_history()

    def _load_order_history(self) -> None:
        orders = self.db.get_orders()
        self.orders_table.setRowCount(0)
        for order in orders[:50]:
            index = self.orders_table.rowCount()
            self.orders_table.insertRow(index)
            vendor_item = QTableWidgetItem(order["vendor"])
            vendor_item.setData(Qt.UserRole, order.get("id"))
            self.orders_table.setItem(index, 0, vendor_item)
            self.orders_table.setItem(index, 1, QTableWidgetItem(order["created_at"]))
            updated_at = order.get("updated_at") or ""
            status = "Edited" if updated_at else "Created"
            self.orders_table.setItem(index, 2, QTableWidgetItem(updated_at))
            self.orders_table.setItem(index, 3, QTableWidgetItem(status))
            self.orders_table.setItem(index, 4, QTableWidgetItem(f"{order['total_amount']:.2f}"))
            self.orders_table.setItem(index, 5, QTableWidgetItem(order["order_filename"] or ""))
        # clear selection and disable action buttons until user picks a row
        self.orders_table.clearSelection()
        self.on_order_selection_changed()

    def on_order_selection_changed(self) -> None:
        sel = self.orders_table.currentRow()
        has = sel is not None and sel >= 0
        self.open_order_btn.setEnabled(has)
        self.export_order_excel_btn.setEnabled(has)
        self.export_order_pdf_btn.setEnabled(has)
        self.delete_order_btn.setEnabled(has)

    def on_vendor_selected(self) -> None:
        selected_items = self.vendor_list.selectedItems()
        if not selected_items:
            return
        self.current_vendor = selected_items[0].text()
        self._load_products_for_vendor(self.current_vendor)

    def _load_products_for_vendor(self, vendor: str) -> None:
        self.product_table.setRowCount(0)
        self.product_cache = self.db.get_vendor_products(vendor)
        for product in self.product_cache:
            index = self.product_table.rowCount()
            self.product_table.insertRow(index)
            display_sku = self._display_sku(product)
            self.product_table.setItem(index, 0, QTableWidgetItem(display_sku))
            self.product_table.setItem(index, 1, QTableWidgetItem(str(product.get("product_name") or "")))
            self.product_table.setItem(index, 2, QTableWidgetItem(str(product.get("brand") or "")))
            self.product_table.setItem(index, 3, QTableWidgetItem(str(product.get("pack") or "")))
            self.product_table.setItem(index, 4, QTableWidgetItem(str(product.get("unit") or "")))
            self.product_table.setItem(index, 5, QTableWidgetItem(str(product.get("ctn_qty") or "0")))
            self.product_table.setItem(index, 6, QTableWidgetItem(f"{product.get('price') or 0.0:.2f}"))
            self.product_table.setItem(index, 7, QTableWidgetItem(str(product.get("stock") or "")))
            self.product_table.setItem(index, 8, QTableWidgetItem(str(product.get("barcode") or "")))
            self.product_table.setItem(index, 9, QTableWidgetItem(str(product.get("sr_sku") or "")))
        if self.current_vendor:
            stats = self.db.get_vendor_statistics(self.current_vendor)
            self.stats_label.setText(
                f"{self.current_vendor}: {stats['product_count']} products, inventory value €{stats['inventory_value']:.2f}"
            )
            self._load_vendor_profile_summary(self.current_vendor)

    def _load_vendor_profile_summary(self, vendor: str) -> None:
        profile = self.db.get_vendor_profile(vendor)
        if not profile:
            self.vendor_detail_summary.setText(
                "No vendor details saved. Click <b>Vendor details</b> to add business info like country of origin, IBAN, tax IDs and contact details."
            )
            return

        display_name = profile.get("display_name") or vendor
        lines = [f"<b>{display_name}</b>"]
        if profile.get("legal_name"):
            lines.append(f"Legal name: {profile.get('legal_name')}")
        if profile.get("customer_number"):
            lines.append(f"Customer #: {profile.get('customer_number')}")
        if profile.get("customer_website"):
            site = profile.get("customer_website")
            lines.append(f"Website: <a href=\"{site}\">{site}</a>")
        if profile.get("country_of_origin"):
            lines.append(f"Country of origin: {profile.get('country_of_origin')}")
        if profile.get("address"):
            lines.append(f"Address: {profile.get('address').replace(chr(10), '<br/>')}")
        if profile.get("contact_name"):
            lines.append(f"Contact: {profile.get('contact_name')}")
        if profile.get("contact_email"):
            lines.append(f"Email: {profile.get('contact_email')}")
        if profile.get("contact_phone"):
            lines.append(f"Phone: {profile.get('contact_phone')}")
        if profile.get("tax_id"):
            lines.append(f"Tax ID: {profile.get('tax_id')}")
        if profile.get("vat_id"):
            lines.append(f"VAT ID: {profile.get('vat_id')}")
        if profile.get("iban"):
            lines.append(f"IBAN: {profile.get('iban')}")
        if profile.get("bank_name"):
            lines.append(f"Bank: {profile.get('bank_name')}")
        if profile.get("swift_bic"):
            lines.append(f"SWIFT/BIC: {profile.get('swift_bic')}")
        self.vendor_detail_summary.setText("<br/>".join(lines))

    def open_vendor_details(self) -> None:
        if not self.current_vendor:
            QMessageBox.warning(self, "No vendor selected", "Please select a vendor before opening vendor details.")
            return
        profile = self.db.get_vendor_profile(self.current_vendor)
        dialog = VendorProfileDialog(self.current_vendor, profile, self.db, parent=self)
        if dialog.exec_():
            self._load_vendor_profile_summary(self.current_vendor)

    def _display_sku(self, product: dict) -> str:
        import json
        sku = product.get("sku") or ""
        pname = product.get("product_name") or ""
        sku_str = str(sku).strip()
        if not sku_str:
            extra = product.get("extra_json") or ""
            try:
                data = json.loads(extra)
            except Exception:
                data = {}
            candidate = None
            for k, v in data.items():
                if "sku" in str(k).lower() and v:
                    candidate = str(v).strip()
                    break
            if not candidate and isinstance(data.get("raw_excel"), dict):
                for k, v in data.get("raw_excel").items():
                    if "sku" in str(k).lower() and v:
                        candidate = str(v).strip()
                        break
            return candidate or ""
        if pname and sku_str.strip() == str(pname).strip():
            extra = product.get("extra_json") or ""
            try:
                data = json.loads(extra)
            except Exception:
                data = {}
            candidate = None
            for k, v in data.items():
                if "sku" in str(k).lower() and v:
                    candidate = str(v).strip()
                    break
            if not candidate and isinstance(data.get("raw_excel"), dict):
                for k, v in data.get("raw_excel").items():
                    if "sku" in str(k).lower() and v:
                        candidate = str(v).strip()
                        break
            return candidate or sku_str
        return sku_str

    def sync_selected_vendor(self) -> None:
        if not self.current_vendor:
            QMessageBox.warning(self, "No vendor selected", "Please select a vendor before syncing.")
            return
        self._sync_vendor(self.current_vendor)

    def reset_selected_vendor_products(self) -> None:
        if not self.current_vendor:
            QMessageBox.warning(self, "No vendor selected", "Please select a vendor before resetting products.")
            return
        response = QMessageBox.question(
            self,
            "Reset vendor products",
            f"This will remove all stored product rows for vendor '{self.current_vendor}' from the database.\nDo you want to continue?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if response != QMessageBox.Yes:
            return

        try:
            self.db.reset_vendor_products(self.current_vendor)
            self._load_products_for_vendor(self.current_vendor)
            self._refresh_all_stats()
            QMessageBox.information(self, "Reset complete", f"Product data for '{self.current_vendor}' has been cleared.")
        except Exception as exc:
            QMessageBox.warning(self, "Reset failed", f"Could not reset vendor product data:\n{exc}")

    def sync_all_vendors(self) -> None:
        for vendor in self.vendor_names:
            self._sync_vendor(vendor)
        self._refresh_all_stats()

    def _sync_vendor(self, vendor: str) -> None:
        products = self.loader.load_products(vendor)
        if not products:
            QMessageBox.warning(self, "Sync failed", f"No product Excel file found for vendor '{vendor}'.")
            return
        self.db.upsert_vendor_products(vendor, products)
        if self.current_vendor == vendor:
            self._load_products_for_vendor(vendor)
        self._refresh_all_stats()
        QMessageBox.information(self, "Sync complete", f"{len(products)} products synced for {vendor}.")

    def open_inventory_dialog(self) -> None:
        if not self.current_vendor:
            QMessageBox.warning(self, "No vendor selected", "Please select a vendor before opening product inventory.")
            return
        dialog = InventoryDialog(self.current_vendor, self.vendor_names, self.db, parent=self)
        dialog.exec_()
        self._load_products_for_vendor(self.current_vendor)
        self._refresh_all_stats()

    def open_order_dialog(self) -> None:
        if not self.current_vendor:
            QMessageBox.warning(self, "No vendor selected", "Please select a vendor before creating an order.")
            return
        if not self.product_cache:
            QMessageBox.warning(self, "No products", "Load vendor products before creating an order.")
            return
        dialog = OrderDialog(self.current_vendor, self.product_cache, self.db, parent=self)
        if dialog.exec_():
            self._load_order_history()

    def _get_selected_order_id(self) -> Optional[int]:
        sel = self.orders_table.currentRow()
        if sel < 0:
            return None
        item = self.orders_table.item(sel, 0)
        if not item:
            return None
        return item.data(Qt.UserRole)

    def on_open_selected_order(self) -> None:
        order_id = self._get_selected_order_id()
        if not order_id:
            QMessageBox.warning(self, "No order selected", "Please select an order to open.")
            return
        # find vendor name from table
        sel = self.orders_table.currentRow()
        vendor = self.orders_table.item(sel, 0).text()
        # load order items
        items = self.db.get_order_items(order_id)
        dialog = OrderDialog(vendor, self.db.get_vendor_products(vendor), self.db, parent=self, order_id=order_id, initial_items=items)
        if dialog.exec_():
            self._load_order_history()

    def on_export_selected_order(self, format: str = "excel") -> None:
        order_id = self._get_selected_order_id()
        if not order_id:
            QMessageBox.warning(self, "No order selected", "Please select an order to export.")
            return
        sel = self.orders_table.currentRow()
        vendor = self.orders_table.item(sel, 0).text()
        items = self.db.get_order_items(order_id)
        dialog = OrderDialog(vendor, self.db.get_vendor_products(vendor), self.db, parent=self, order_id=order_id, initial_items=items)
        if format == "excel":
            dialog._export_to_excel()
        else:
            dialog._export_to_pdf()
        self._load_order_history()

    def on_delete_selected_order(self) -> None:
        order_id = self._get_selected_order_id()
        if not order_id:
            QMessageBox.warning(self, "No order selected", "Please select an order to delete.")
            return
        resp = QMessageBox.question(self, "Delete order", "Are you sure you want to delete the selected order?", QMessageBox.Yes | QMessageBox.No)
        if resp != QMessageBox.Yes:
            return
        try:
            self.db.delete_order(order_id)
            QMessageBox.information(self, "Deleted", "Order deleted")
            self._load_order_history()
        except Exception as exc:
            QMessageBox.warning(self, "Delete failed", f"Could not delete order:\n{exc}")

    def on_create_demo_order(self) -> None:
        # create a small demo order to populate history for testing
        vendor = self.current_vendor or (self.vendor_names[0] if self.vendor_names else "demo_vendor")
        items = [
            {"sku": "DEMO-001", "product_name": "Sample Milk 1L", "brand": "DemoBrand", "quantity": 10, "ctn_qty": 1, "unit_price": 0.95, "total_price": 9.5},
            {"sku": "DEMO-002", "product_name": "Sample Bread", "brand": "DemoBake", "quantity": 5, "ctn_qty": 1, "unit_price": 1.2, "total_price": 6.0},
        ]
        total = sum(i.get("total_price", 0.0) for i in items)
        try:
            order_id = self.db.save_order(vendor=vendor, items=items, total_amount=total, order_filename="", notes="Demo order")
            QMessageBox.information(self, "Demo order", f"Demo order created with id {order_id}")
        except Exception as exc:
            QMessageBox.warning(self, "Demo failed", f"Could not create demo order:\n{exc}")
        self._load_order_history()


def main() -> None:
    app = QApplication(sys.argv)
    window = VendorManagerApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
