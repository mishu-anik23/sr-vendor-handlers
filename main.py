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
    QWidget,
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
        self.setWindowTitle(f"Create Order - {vendor}")
        self.setMinimumWidth(540)
        self.selected_item: Optional[Dict] = None
        self._build_ui()
        self._load_products()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        self.vendor_label = QLabel(f"Vendor: <b>{self.vendor}</b>")
        self.product_combo = QComboBox()
        self.brand_input = QLineEdit()
        self.quantity_input = QLineEdit("1")
        self.ctn_qty_input = QLineEdit("1")
        self.unit_price_input = QLineEdit("0.00")
        self.notes_input = QTextEdit()
        self.notes_input.setPlaceholderText("Order notes or instructions")
        self.total_label = QLabel("Total: 0.00")
        self.order_button = QPushButton("Generate Order Sheet")
        self.order_button.clicked.connect(self.on_generate_order)

        self.product_combo.currentIndexChanged.connect(self.on_product_changed)
        self.quantity_input.textChanged.connect(self._update_total)
        self.ctn_qty_input.textChanged.connect(self._update_total)
        self.unit_price_input.textChanged.connect(self._update_total)

        grid = QGridLayout()
        grid.addWidget(self.vendor_label, 0, 0, 1, 2)
        grid.addWidget(QLabel("Product:"), 1, 0)
        grid.addWidget(self.product_combo, 1, 1)
        grid.addWidget(QLabel("Brand:"), 2, 0)
        grid.addWidget(self.brand_input, 2, 1)
        grid.addWidget(QLabel("Quantity:"), 3, 0)
        grid.addWidget(self.quantity_input, 3, 1)
        grid.addWidget(QLabel("CTN Qty:"), 4, 0)
        grid.addWidget(self.ctn_qty_input, 4, 1)
        grid.addWidget(QLabel("Unit Price:"), 5, 0)
        grid.addWidget(self.unit_price_input, 5, 1)
        grid.addWidget(QLabel("Notes:"), 6, 0)
        grid.addWidget(self.notes_input, 6, 1)
        grid.addWidget(self.total_label, 7, 0, 1, 2)
        grid.addWidget(self.order_button, 8, 0, 1, 2)

        layout.addLayout(grid)

    def _load_products(self) -> None:
        self.product_combo.clear()
        for product in self.products:
            name = product.get("product_name") or "Unnamed product"
            self.product_combo.addItem(name, product)
        if self.products:
            self.product_combo.setCurrentIndex(0)
            self.on_product_changed(0)

    def on_product_changed(self, index: int) -> None:
        product = self.product_combo.itemData(index)
        if not product:
            return
        self.selected_item = product
        self.brand_input.setText(str(product.get("brand") or ""))
        self.quantity_input.setText(str(product.get("stock") or "1"))
        self.ctn_qty_input.setText(str(product.get("ctn_qty") or "1"))
        self.unit_price_input.setText(f"{product.get('price') or 0.0:.2f}")
        self._update_total()

    def _update_total(self) -> None:
        try:
            quantity = int(self.quantity_input.text() or 0)
        except ValueError:
            quantity = 0
        try:
            ctn_qty = int(self.ctn_qty_input.text() or 0)
        except ValueError:
            ctn_qty = 0
        try:
            unit_price = float(self.unit_price_input.text() or 0.0)
        except ValueError:
            unit_price = 0.0
        total = (quantity * ctn_qty) * unit_price
        self.total_label.setText(f"Total: {total:.2f}")

    def on_generate_order(self) -> None:
        if self.selected_item is None:
            QMessageBox.warning(self, "No product", "Please select a product to order.")
            return
        try:
            quantity = int(self.quantity_input.text() or 0)
            ctn_qty = int(self.ctn_qty_input.text() or 0)
            unit_price = float(self.unit_price_input.text() or 0.0)
        except ValueError:
            QMessageBox.warning(self, "Invalid input", "Please enter valid numeric values for quantity, CTN quantity, and unit price.")
            return
        order_item = {
            "source_id": self.selected_item.get("source_id"),
            "product_name": self.selected_item.get("product_name"),
            "brand": self.brand_input.text().strip() or self.selected_item.get("brand"),
            "quantity": quantity,
            "ctn_qty": ctn_qty,
            "unit_price": unit_price,
            "total_price": quantity * ctn_qty * unit_price,
            "pack": self.selected_item.get("pack"),
            "raw_json": self.selected_item.get("extra_json"),
        }
        order_rows = [order_item]
        filename = f"{self.vendor}_order_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        order_dir = VENDOR_ROOT / self.vendor / "orders"
        order_dir.mkdir(parents=True, exist_ok=True)
        order_path = order_dir / filename
        try:
            from pandas import DataFrame

            DataFrame(
                [
                    {
                        "Vendor": self.vendor,
                        "Product": order_item["product_name"],
                        "Brand": order_item["brand"],
                        "Quantity": order_item["quantity"],
                        "CTN Qty": order_item["ctn_qty"],
                        "Unit Price": order_item["unit_price"],
                        "Total Price": order_item["total_price"],
                    }
                ]
            ).to_excel(order_path, index=False)
        except Exception as exc:
            QMessageBox.critical(self, "Order creation failed", f"Could not create order Excel sheet:\n{exc}")
            return
        self.db.save_order(
            vendor=self.vendor,
            items=[order_item],
            total_amount=order_item["total_price"],
            order_filename=str(order_path),
            notes=self.notes_input.toPlainText().strip(),
        )
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
        self.stats_table = QTableWidget(0, 3)
        self.stats_table.setHorizontalHeaderLabels(["Vendor", "Products", "Inventory value"])
        self.stats_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        stats_layout.addWidget(self.stats_label)
        stats_layout.addWidget(self.stats_table)
        stats_box.setLayout(stats_layout)
        right_panel.addWidget(stats_box)

        self.product_table = QTableWidget(0, 8)
        self.product_table.setHorizontalHeaderLabels(
            ["SKU", "Product", "Brand", "Pack", "Unit", "CTN Qty", "Price", "Stock"]
        )
        self.product_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
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
            self.product_table.setItem(index, 0, QTableWidgetItem(str(product.get("source_id") or "")))
            self.product_table.setItem(index, 1, QTableWidgetItem(str(product.get("product_name") or "")))
            self.product_table.setItem(index, 2, QTableWidgetItem(str(product.get("brand") or "")))
            self.product_table.setItem(index, 3, QTableWidgetItem(str(product.get("pack") or "")))
            self.product_table.setItem(index, 4, QTableWidgetItem(str(product.get("unit") or "")))
            self.product_table.setItem(index, 5, QTableWidgetItem(str(product.get("ctn_qty") or "0")))
            self.product_table.setItem(index, 6, QTableWidgetItem(f"{product.get('price') or 0.0:.2f}"))
            self.product_table.setItem(index, 7, QTableWidgetItem(str(product.get("stock") or "")))
        if self.current_vendor:
            stats = self.db.get_vendor_statistics(self.current_vendor)
            self.stats_label.setText(
                f"{self.current_vendor}: {stats['product_count']} products, inventory value €{stats['inventory_value']:.2f}"
            )

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
