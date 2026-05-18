import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
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
    def __init__(self, vendor: str, products: List[Dict], db: DatabaseManager, parent=None):
        super().__init__(parent)
        self.vendor = vendor
        self.products = products
        self.db = db
        self.cart_items: List[Dict] = []  # Store items added to cart
        self.setWindowTitle(f"Create Order - {vendor}")
        self.setMinimumSize(1200, 800)
        self._build_ui()
        self._load_products_table()

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
        layout.addWidget(self.products_table)
        
        # Cart section
        layout.addWidget(QLabel("<b>Order Cart</b>"))
        self.cart_table = QTableWidget(0, 9)
        self.cart_table.setHorizontalHeaderLabels(
            ["SKU", "Brand", "Product", "Pack", "Qty", "CTN Qty", "Unit Price", "Total", "Remove"]
        )
        header = self.cart_table.horizontalHeader()
        for col in range(self.cart_table.columnCount()):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        self.cart_table.setColumnWidth(0, 140)
        self.cart_table.setColumnWidth(1, 120)
        self.cart_table.setColumnWidth(2, 420)
        self.cart_table.setColumnWidth(3, 120)
        self.cart_table.setColumnWidth(4, 80)
        self.cart_table.setColumnWidth(5, 90)
        self.cart_table.setColumnWidth(6, 100)
        self.cart_table.setColumnWidth(7, 120)
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
        self.generate_button = QPushButton("Generate Order Sheet")
        self.generate_button.clicked.connect(self.on_generate_order)
        button_row.addWidget(self.generate_button)
        
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
            self.products_table.setCellWidget(row, 6, qty_combo)
            
            # Custom qty input
            custom_qty = QSpinBox()
            custom_qty.setMinimum(0)
            custom_qty.setMaximum(10000)
            custom_qty.setValue(0)
            self.products_table.setCellWidget(row, 7, custom_qty)
            
            # Add to cart button
            add_btn = QPushButton("Add")
            add_btn.clicked.connect(lambda checked, r=row: self.on_add_to_cart(r))
            self.products_table.setCellWidget(row, 8, add_btn)

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

    def on_add_to_cart(self, row: int) -> None:
        product = self.product_row_map.get(row)
        if not product:
            return
        
        qty_combo = self.products_table.cellWidget(row, 6)
        custom_qty = self.products_table.cellWidget(row, 7)
        
        qty_0_10 = int(qty_combo.currentText() or 0)
        custom_qty_val = custom_qty.value() if custom_qty else 0
        
        # Use custom qty if > 10, otherwise use combo selection
        quantity = custom_qty_val if custom_qty_val > 10 else qty_0_10
        
        if quantity <= 0:
            QMessageBox.warning(self, "Invalid quantity", "Please select a quantity > 0")
            return
        
        # Add to cart
        cart_item = {
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
        self.cart_items.append(cart_item)
        self._update_cart_table()
        
        # Reset inputs
        qty_combo.setCurrentText("0")
        custom_qty.setValue(0)
        QMessageBox.information(self, "Added", f"Added {quantity} unit(s) to cart")

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
            
            # Pack
            self.cart_table.setItem(row, 3, QTableWidgetItem(str(item.get("pack") or "")))
            
            # Qty
            self.cart_table.setItem(row, 4, QTableWidgetItem(str(item.get("quantity") or 0)))
            
            # CTN Qty
            self.cart_table.setItem(row, 5, QTableWidgetItem(str(item.get("ctn_qty") or 0)))
            
            # Unit Price
            price = float(item.get("unit_price") or 0.0)
            self.cart_table.setItem(row, 6, QTableWidgetItem(f"{price:.2f}"))
            
            # Total
            total = float(item.get("total_price") or 0.0)
            self.cart_table.setItem(row, 6, QTableWidgetItem(f"{total:.2f}"))
            total_amount += total
            
            # Remove button
            remove_btn = QPushButton("Remove")
            remove_btn.clicked.connect(lambda checked, r=row: self.on_remove_from_cart(r))
            self.cart_table.setCellWidget(row, 7, remove_btn)
        
        self.total_label.setText(f"Cart Total: EUR {total_amount:.2f}")

    def on_remove_from_cart(self, row: int) -> None:
        if 0 <= row < len(self.cart_items):
            self.cart_items.pop(row)
            self._update_cart_table()

    def on_generate_order(self) -> None:
        if not self.cart_items:
            QMessageBox.warning(self, "Empty cart", "Please add items to cart before generating order.")
            return
        
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
                    "Pack": item.get("pack"),
                    "Unit": item.get("unit"),
                    "Quantity": item.get("quantity"),
                    "CTN Qty": item.get("ctn_qty"),
                    "Unit Price": item.get("unit_price"),
                    "Total Price": item.get("total_price"),
                })
            
            DataFrame(order_data).to_excel(order_path, index=False)
        except Exception as exc:
            QMessageBox.critical(self, "Order creation failed", f"Could not create order Excel sheet:\n{exc}")
            return
        
        # Save to database
        try:
            self.db.save_order(
                vendor=self.vendor,
                items=self.cart_items,
                total_amount=sum(item.get("total_price", 0.0) for item in self.cart_items),
                order_filename=str(order_path),
                notes=self.notes_input.toPlainText().strip(),
            )
        except Exception as exc:
            QMessageBox.warning(self, "Database save", f"Order Excel saved but DB save failed:\n{exc}")
        
        QMessageBox.information(self, "Order saved", f"Order sheet created:\n{order_path}")
        self.accept()


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
        self.order_button = QPushButton("Create order")
        self.order_button.clicked.connect(self.open_order_dialog)
        left_panel.addWidget(self.sync_vendor_button)
        left_panel.addWidget(self.sync_all_button)
        left_panel.addWidget(self.order_button)

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
        self.orders_table = QTableWidget(0, 4)
        self.orders_table.setHorizontalHeaderLabels(["Vendor", "Date", "Total", "Filename"])
        self.orders_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
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
            self.orders_table.setItem(index, 0, QTableWidgetItem(order["vendor"]))
            self.orders_table.setItem(index, 1, QTableWidgetItem(order["created_at"]))
            self.orders_table.setItem(index, 2, QTableWidgetItem(f"{order['total_amount']:.2f}"))
            self.orders_table.setItem(index, 3, QTableWidgetItem(order["order_filename"] or ""))

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


def main() -> None:
    app = QApplication(sys.argv)
    window = VendorManagerApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
