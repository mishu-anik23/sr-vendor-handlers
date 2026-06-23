import sys
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import io
import sqlite3

import pandas as pd
import requests
from PyQt5.QtCore import Qt, QEvent
from PyQt5.QtGui import QIcon, QPixmap, QImage, QStandardItemModel, QStandardItem
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
    QListWidget, QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget, QHBoxLayout, QSpinBox, QScrollArea,
    QFileDialog,
    QRadioButton, QButtonGroup, QInputDialog,
    QStylePainter, QStyleOptionComboBox, QStyle, QListView,
)

from db_manager import DatabaseManager
from excel_loader import ExcelLoader
import io
import pypdfium2 as pdfium
from PIL import Image
from invoice_parsers import get_parser


COMPANY_NAME = "Sunrise Supermarket"
COMPANY_ADDRESS = "Schwarzwald Straße 27, 60528 Frankfurt"
LOGO_FILE = Path(__file__).resolve().parent / "logo-sr-tmp.jpeg"
DB_FILE = Path(__file__).resolve().parent / "data" / "vendor_app.db"
VENDOR_ROOT = Path(__file__).resolve().parent / "data" / "vendors"


class CheckableComboBox(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setView(QListView(self))
        self.view().pressed.connect(self.handle_item_pressed)
        self.setModel(QStandardItemModel(self))
        self.view().viewport().installEventFilter(self)
        self._changed = False

    def eventFilter(self, widget, event):
        if widget == self.view().viewport() and event.type() == QEvent.MouseButtonRelease:
            index = self.view().indexAt(event.pos())
            item = self.model().itemFromIndex(index)
            if item is not None:
                item.setCheckState(Qt.Checked if item.checkState() == Qt.Unchecked else Qt.Unchecked)
            self._changed = True
            return True
        return super().eventFilter(widget, event)

    def handle_item_pressed(self, index):
        self._changed = True

    def hidePopup(self):
        if self._changed:
            self._changed = False
            self.currentTextChanged.emit(self.currentText())
        super().hidePopup()

    def add_items(self, items):
        self.clear()
        for text in items:
            item = QStandardItem(text)
            item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            item.setData(Qt.Unchecked, Qt.CheckStateRole)
            self.model().appendRow(item)

    def checked_items(self):
        checked = []
        for i in range(self.model().rowCount()):
            item = self.model().item(i)
            if item and item.checkState() == Qt.Checked:
                checked.append(item.text())
        return checked

    def paintEvent(self, event):
        painter = QStylePainter(self)
        painter.setFont(self.font())
        opt = QStyleOptionComboBox()
        self.initStyleOption(opt)
        checked = self.checked_items()
        if not checked:
            opt.currentText = "All"
        else:
            opt.currentText = ", ".join(checked)
        painter.drawComplexControl(QStyle.CC_ComboBox, opt)
        painter.drawControl(QStyle.CE_ComboBoxLabel, opt)


class TagListDialog(QDialog):
    def __init__(self, selected_rows, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Tag List Manager")
        self.setMinimumSize(1000, 700)
        self.selected_rows = selected_rows  # List of dicts
        self.generated_pdf_data = None
        self.logo_path = Path(__file__).resolve().parent / "logo-sr-tmp.jpeg"
        self._build_ui()
        self.load_saved_tags()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        self.header_label = QLabel(f"Selected Rows for Tag Generation ({len(self.selected_rows)} selected)")
        self.header_label.setStyleSheet("font-size: 12pt; font-weight: bold;")
        layout.addWidget(self.header_label)

        self.table = QTableWidget()
        self.columns = []
        if self.selected_rows:
            self.columns = list(self.selected_rows[0].keys())
        self.table.setColumnCount(len(self.columns))
        self.table.setHorizontalHeaderLabels(self.columns)
        self.table.setRowCount(len(self.selected_rows))

        for r_idx, row_data in enumerate(self.selected_rows):
            for c_idx, col_name in enumerate(self.columns):
                val = str(row_data.get(col_name, ''))
                item = QTableWidgetItem(val)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # non-editable
                self.table.setItem(r_idx, c_idx, item)

        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        layout.addWidget(self.table)

        controls_layout = QHBoxLayout()

        self.process_btn = QPushButton("Process Tag Entries")
        self.process_btn.setStyleSheet("background-color: #ffc107; color: black; font-weight: bold; padding: 6px;")
        self.process_btn.clicked.connect(self.process_entries)
        controls_layout.addWidget(self.process_btn)

        self.save_btn = QPushButton("Save")
        self.save_btn.setStyleSheet("background-color: #28a745; color: white; font-weight: bold; padding: 6px;")
        self.save_btn.clicked.connect(self.save_entries)
        controls_layout.addWidget(self.save_btn)

        self.generate_btn = QPushButton("Generate Tag")
        self.generate_btn.setStyleSheet("background-color: #17a2b8; color: white; font-weight: bold; padding: 6px;")
        self.generate_btn.clicked.connect(self.generate_tags_pdf)
        controls_layout.addWidget(self.generate_btn)

        self.download_pdf_btn = QPushButton("Download Generated Tags")
        self.download_pdf_btn.setStyleSheet("background-color: #6c757d; color: white; font-weight: bold; padding: 6px;")
        self.download_pdf_btn.setEnabled(False)
        self.download_pdf_btn.clicked.connect(self.download_tags_pdf)
        controls_layout.addWidget(self.download_pdf_btn)

        layout.addLayout(controls_layout)

        self.preview_tab_widget = QTabWidget()

        self.db_tab = QWidget()
        db_tab_layout = QVBoxLayout(self.db_tab)
        self.db_table = QTableWidget()
        db_tab_layout.addWidget(self.db_table)
        self.preview_tab_widget.addTab(self.db_tab, "Saved Tag Rows")

        self.pdf_tab = QWidget()
        pdf_tab_layout = QVBoxLayout(self.pdf_tab)
        self.pdf_scroll = QScrollArea()
        self.pdf_scroll.setWidgetResizable(True)
        self.pdf_preview_container = QWidget()
        self.pdf_preview_layout = QVBoxLayout(self.pdf_preview_container)
        self.pdf_scroll.setWidget(self.pdf_preview_container)
        pdf_tab_layout.addWidget(self.pdf_scroll)
        self.preview_tab_widget.addTab(self.pdf_tab, "Generated PDF Preview")

        layout.addWidget(self.preview_tab_widget)

    def process_entries(self):
        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                if item:
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
        QMessageBox.information(self, "Editable Mode", "Table entries are now editable.")

    def save_entries(self):
        rows_data = []
        for r in range(self.table.rowCount()):
            row_vals = []
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                row_vals.append(item.text().strip() if item else "")
            rows_data.append(row_vals)

        if not rows_data:
            QMessageBox.warning(self, "No Data", "No rows to save.")
            return

        try:
            conn = sqlite3.connect("sr_product_tags.db")
            cursor = conn.cursor()
            col_defs = ", ".join([f"`{col}` TEXT" for col in self.columns])
            create_sql = f"CREATE TABLE IF NOT EXISTS product_tags (id INTEGER PRIMARY KEY AUTOINCREMENT, {col_defs}, saved_time TEXT)"
            cursor.execute(create_sql)

            saved_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            placeholders = ", ".join(["?"] * (len(self.columns) + 1))
            insert_cols = ", ".join([f"`{col}`" for col in self.columns]) + ", saved_time"
            insert_sql = f"INSERT INTO product_tags ({insert_cols}) VALUES ({placeholders})"

            for row in rows_data:
                val_list = list(row) + [saved_time]
                cursor.execute(insert_sql, val_list)

            conn.commit()
            conn.close()

            QMessageBox.information(self, "Success", "Selected entries saved to 'sr_product_tags' SQLite DB successfully.")
            self.load_saved_tags()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save entries to SQLite:\n{str(e)}")

    def load_saved_tags(self):
        try:
            conn = sqlite3.connect("sr_product_tags.db")
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='product_tags'")
            if not cursor.fetchone():
                conn.close()
                self.db_table.setColumnCount(0)
                self.db_table.setRowCount(0)
                return

            cursor.execute("SELECT * FROM product_tags ORDER BY id DESC")
            rows = cursor.fetchall()
            colnames = [description[0] for description in cursor.description]
            conn.close()

            self.db_table.setColumnCount(len(colnames))
            self.db_table.setHorizontalHeaderLabels(colnames)
            self.db_table.setRowCount(len(rows))

            for r_idx, row in enumerate(rows):
                for c_idx, val in enumerate(row):
                    self.db_table.setItem(r_idx, c_idx, QTableWidgetItem(str(val) if val is not None else ""))
            self.db_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        except Exception as e:
            print(f"Error loading saved tags: {e}")

    def generate_tags_pdf(self):
        products = []
        for r in range(self.table.rowCount()):
            row_data = {}
            for c_idx, col_name in enumerate(self.columns):
                item = self.table.item(r, c_idx)
                row_data[col_name] = item.text().strip() if item else ""
            products.append(row_data)

        if not products:
            QMessageBox.warning(self, "No Products", "No product rows found to generate tags.")
            return

        try:
            import tempfile
            import os
            temp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            temp_pdf.close()
            pdf_path = temp_pdf.name

            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import A4
            from reportlab.graphics.barcode import createBarcodeDrawing

            c = canvas.Canvas(pdf_path, pagesize=A4)

            tag_width = 180
            tag_height = 95
            margin_x = 20
            margin_y = 30
            gap_x = 10
            gap_y = 10

            total_tags = len(products)
            tags_per_page = 24
            date_str = datetime.now().strftime("%Y-%m-%d")

            for idx, product in enumerate(products):
                page_idx = idx % tags_per_page
                if idx > 0 and page_idx == 0:
                    c.showPage()

                col = page_idx % 3
                row = page_idx // 3

                x = margin_x + col * (tag_width + gap_x)
                y = 841.89 - margin_y - (row + 1) * tag_height - row * gap_y

                c.setStrokeColorRGB(0.7, 0.7, 0.7)
                c.setLineWidth(0.5)
                c.rect(x, y, tag_width, tag_height)

                if self.logo_path.exists():
                    try:
                        c.drawImage(str(self.logo_path), x + 5, y + tag_height - 25, width=35, height=20, preserveAspectRatio=True)
                    except Exception as e:
                        print(f"Error drawing logo in PDF: {e}")

                c.setFont("Helvetica", 6)
                c.setFillColorRGB(0.4, 0.4, 0.4)
                c.drawRightString(x + tag_width - 5, y + tag_height - 12, f"Date: {date_str}")

                c.setFont("Helvetica-Bold", 8)
                c.setFillColorRGB(0, 0, 0)
                name_str = product.get('Name', product.get('product_name', product.get('Description', '')))

                words = name_str.split()
                lines = []
                current_line = ""
                for w in words:
                    if c.stringWidth(current_line + " " + w, "Helvetica-Bold", 8) < 170:
                        current_line += (" " if current_line else "") + w
                    else:
                        lines.append(current_line)
                        current_line = w
                if current_line:
                    lines.append(current_line)

                y_offset = y + tag_height - 35
                for line in lines[:2]:
                    c.drawString(x + 5, y_offset, line)
                    y_offset -= 10

                price_val = ""
                for k in ['sale_price', 'Sale price', 'sale price', 'Price']:
                    if k in product:
                        price_val = product[k]
                        break
                if not price_val:
                    for col_name in product.keys():
                        if 'price' in col_name.lower():
                            price_val = product[col_name]
                            break

                try:
                    cleaned_price = re.sub(r'[^\d.]', '', str(price_val))
                    if cleaned_price:
                        price_display = f"€ {float(cleaned_price):.2f}"
                    else:
                        price_display = str(price_val)
                except Exception:
                    price_display = str(price_val)

                c.setFont("Helvetica-Bold", 20)
                c.drawRightString(x + tag_width - 5, y + 25, price_display)

                barcode_val = ""
                for k in ['Barcode', 'barcode', 'APC', 'Art No', 'Item No.']:
                    if k in product and product[k] and str(product[k]).lower() != 'nan':
                        barcode_val = str(product[k]).strip()
                        break

                if barcode_val:
                    digits_only = re.sub(r'\D', '', barcode_val)
                    drawing = None
                    for b_type in ['EAN13', 'Code128']:
                        if b_type == 'EAN13' and len(digits_only) != 13:
                            continue
                        try:
                            drawing = createBarcodeDrawing(b_type, value=digits_only if b_type == 'EAN13' else barcode_val,
                                                           barHeight=20, barWidth=0.8, humanReadable=True)
                            break
                        except Exception:
                            continue
                    if drawing:
                        drawing.drawOn(c, x + 5, y + 5)
                    else:
                        c.setFont("Helvetica", 7)
                        c.drawString(x + 5, y + 8, f"BC: {barcode_val}")

            c.showPage()
            c.save()

            with open(pdf_path, "rb") as f:
                self.generated_pdf_data = f.read()

            try:
                os.unlink(pdf_path)
            except Exception:
                pass

            self.download_pdf_btn.setEnabled(True)
            QMessageBox.information(self, "Success", "Supermarket tags generated successfully. Check the PDF Preview tab.")

            self.preview_tab_widget.setCurrentIndex(1)
            self.render_pdf_preview()

        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Error", f"Failed to generate tags PDF:\n{str(e)}")

    def render_pdf_preview(self):
        import io
        while self.pdf_preview_layout.count():
            item = self.pdf_preview_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self.generated_pdf_data:
            return

        try:
            import pypdfium2 as pdfium
            doc = pdfium.PdfDocument(self.generated_pdf_data)
            for page_idx in range(len(doc)):
                page = doc[page_idx]
                bitmap = page.render(scale=2)
                pil_img = bitmap.to_pil()

                im_data = io.BytesIO()
                pil_img.save(im_data, format='PNG')
                qimg = QImage.fromData(im_data.getvalue())

                lbl = QLabel()
                lbl.setPixmap(QPixmap.fromImage(qimg))
                lbl.setAlignment(Qt.AlignCenter)
                lbl.setStyleSheet("border: 1px solid #ccc; margin: 10px; background-color: white;")
                self.pdf_preview_layout.addWidget(lbl)
        except Exception as e:
            print(f"Error rendering PDF preview: {e}")
            lbl = QLabel(f"Failed to render preview:\n{str(e)}")
            self.pdf_preview_layout.addWidget(lbl)

    def download_tags_pdf(self):
        if not self.generated_pdf_data:
            QMessageBox.warning(self, "Error", "No generated PDF data to save.")
            return

        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Tags PDF", "sunrise_tags.pdf", "PDF Files (*.pdf)", options=options)
        if file_name:
            try:
                with open(file_name, "wb") as f:
                    f.write(self.generated_pdf_data)
                QMessageBox.information(self, "Success", "Tags PDF saved successfully.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save PDF file:\n{str(e)}")


def init_brand_db():
    db_path = Path("sr_retail_brands.db")
    conn = sqlite3.connect(str(db_path))
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS brands (
            id INTEGER PRIMARY KEY,
            name TEXT UNIQUE NOT NULL,
            logo_path TEXT,
            origin TEXT
        )
    """)
    conn.commit()
    
    cursor.execute("SELECT COUNT(*) FROM brands")
    count = cursor.fetchone()[0]
    if count == 0:
        try:
            import sys
            from importlib.machinery import SourceFileLoader
            sku_gen_path = Path(__file__).resolve().parent / "sr-sku-gen.py"
            if sku_gen_path.exists():
                sku_gen_mod = SourceFileLoader("sr_sku_gen", str(sku_gen_path)).load_module()
                brand_map = getattr(sku_gen_mod, "BRAND_MAP", {})
                for k, v in brand_map.items():
                    try:
                        b_id = int(k)
                        b_name = str(v).strip().upper()
                        cursor.execute("INSERT OR IGNORE INTO brands (id, name) VALUES (?, ?)", (b_id, b_name))
                    except Exception as ex:
                        print(f"Error inserting brand {k}: {ex}")
                conn.commit()
        except Exception as e:
            print(f"Error loading BRAND_MAP dynamically: {e}")
    conn.close()


def normalize_description_to_name(desc_str: str) -> str:
    if not desc_str:
        return ""
    desc_str = str(desc_str).strip()
    pattern = r'\b\d+\s*[xX*]\s*(\d+(?:\.\d+)?)\s*(kg|g|l|ml|pcs|piece|pieces|oz|lbs)\b'
    
    def repl(match):
        val = match.group(1)
        unit = match.group(2).upper()
        if unit.endswith('S'):
            unit = unit[:-1]
        return f"{val} {unit}"
        
    normalized = re.sub(pattern, repl, desc_str, flags=re.IGNORECASE)
    return normalized


def extract_brand_from_desc(desc_str: str, brands_list: list) -> str:
    if not desc_str:
        return ""
    desc_upper = str(desc_str).upper()
    for _, b_name, _, _ in brands_list:
        b_name_upper = str(b_name).strip().upper()
        pattern = r'\b' + re.escape(b_name_upper) + r'\b'
        if re.search(pattern, desc_upper):
            return b_name_upper
    return ""


class ShopfloorBrandsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Shopfloor Retail Brands")
        self.setMinimumSize(950, 600)
        self.db_path = Path("sr_retail_brands.db")
        self._build_ui()
        self.load_brands()

    def _build_ui(self):
        layout = QHBoxLayout(self)

        # Left Column - Brands Table
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        left_layout.addWidget(QLabel("<b>Brands List (from DB)</b>"))
        
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["ID", "Brand Name", "Logo Path", "Origin"])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.itemSelectionChanged.connect(self.on_row_selected)
        left_layout.addWidget(self.table)
        
        layout.addWidget(left_widget, 2)

        # Right Column - Editor Form
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)

        editor_group = QGroupBox("Brand Details Editor")
        form_layout = QFormLayout(editor_group)

        self.id_label = QLabel("Auto-Generated (New)")
        self.id_label.setStyleSheet("font-weight: bold; color: #007bff;")
        form_layout.addRow("Brand ID:", self.id_label)

        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("e.g. AASHIRVAAD")
        form_layout.addRow("Brand Name:", self.name_input)

        logo_layout = QHBoxLayout()
        self.logo_path_input = QLineEdit()
        self.logo_path_input.setPlaceholderText("logo_path/file.png")
        logo_layout.addWidget(self.logo_path_input)
        
        self.browse_btn = QPushButton("Browse...")
        self.browse_btn.clicked.connect(self.browse_logo)
        logo_layout.addWidget(self.browse_btn)
        form_layout.addRow("Logo Path:", logo_layout)

        self.origin_input = QLineEdit()
        self.origin_input.setPlaceholderText("e.g. India (Optional)")
        form_layout.addRow("Origin:", self.origin_input)

        right_layout.addWidget(editor_group)

        # Action Buttons
        btn_layout = QGridLayout()
        
        self.new_btn = QPushButton("Add New Brand")
        self.new_btn.setStyleSheet("background-color: #17a2b8; color: white; font-weight: bold;")
        self.new_btn.clicked.connect(self.prepare_for_new)
        btn_layout.addWidget(self.new_btn, 0, 0)

        self.save_btn = QPushButton("Save Brand")
        self.save_btn.setStyleSheet("background-color: #28a745; color: white; font-weight: bold;")
        self.save_btn.clicked.connect(self.save_brand)
        btn_layout.addWidget(self.save_btn, 0, 1)

        self.delete_btn = QPushButton("Delete Selected")
        self.delete_btn.setStyleSheet("background-color: #dc3545; color: white;")
        self.delete_btn.clicked.connect(self.delete_brand)
        btn_layout.addWidget(self.delete_btn, 1, 0, 1, 2)

        right_layout.addLayout(btn_layout)
        right_layout.addStretch()

        layout.addWidget(right_widget, 1)

    def load_brands(self):
        try:
            conn = sqlite3.connect(str(self.db_path))
            cursor = conn.cursor()
            cursor.execute("SELECT id, name, logo_path, origin FROM brands ORDER BY id ASC")
            rows = cursor.fetchall()
            conn.close()

            self.table.setRowCount(len(rows))
            for r_idx, row in enumerate(rows):
                for c_idx, val in enumerate(row):
                    item = QTableWidgetItem(str(val) if val is not None else "")
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.table.setItem(r_idx, c_idx, item)
            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        except Exception as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load brands:\n{e}")

    def on_row_selected(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        row = selected_items[0].row()
        b_id = self.table.item(row, 0).text()
        name = self.table.item(row, 1).text()
        logo = self.table.item(row, 2).text()
        origin = self.table.item(row, 3).text()

        self.id_label.setText(b_id)
        self.name_input.setText(name)
        self.logo_path_input.setText(logo)
        self.origin_input.setText(origin)

    def prepare_for_new(self):
        self.table.clearSelection()
        self.id_label.setText("Auto-Generated (New)")
        self.name_input.clear()
        self.logo_path_input.clear()
        self.origin_input.clear()
        self.name_input.setFocus()

    def browse_logo(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Brand Logo", "", "Images (*.png *.jpg *.jpeg *.gif *.svg *.bmp)")
        if file_name:
            import shutil
            logos_dir = Path("data/brand_logos")
            logos_dir.mkdir(parents=True, exist_ok=True)
            dest_path = logos_dir / Path(file_name).name
            try:
                shutil.copy2(file_name, dest_path)
                self.logo_path_input.setText(str(dest_path.relative_to(Path("."))))
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to copy logo:\n{e}")

    def save_brand(self):
        name = self.name_input.text().strip().upper()
        if not name:
            QMessageBox.warning(self, "Input Error", "Brand name is required.")
            return

        logo = self.logo_path_input.text().strip()
        origin = self.origin_input.text().strip()
        is_new = (self.id_label.text() == "Auto-Generated (New)")

        try:
            conn = sqlite3.connect(str(self.db_path))
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS brands (
                    id INTEGER PRIMARY KEY,
                    name TEXT UNIQUE NOT NULL,
                    logo_path TEXT,
                    origin TEXT
                )
            """)

            if is_new:
                cursor.execute("SELECT MAX(id) FROM brands")
                max_val = cursor.fetchone()[0]
                next_id = (max_val or 0) + 1
                cursor.execute(
                    "INSERT INTO brands (id, name, logo_path, origin) VALUES (?, ?, ?, ?)",
                    (next_id, name, logo or None, origin or None)
                )
            else:
                b_id = int(self.id_label.text())
                cursor.execute(
                    "UPDATE brands SET name=?, logo_path=?, origin=? WHERE id=?",
                    (name, logo or None, origin or None, b_id)
                )
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Success", "Brand saved successfully.")
            self.load_brands()
            self.prepare_for_new()
        except sqlite3.IntegrityError:
            QMessageBox.critical(self, "Error", f"Brand name '{name}' already exists in database.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save brand:\n{e}")

    def delete_brand(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Selection Error", "Please select a brand to delete.")
            return
        row = selected_items[0].row()
        b_id = int(self.table.item(row, 0).text())
        b_name = self.table.item(row, 1).text()

        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete brand '{b_name}' (ID: {b_id})?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            try:
                conn = sqlite3.connect(str(self.db_path))
                cursor = conn.cursor()
                cursor.execute("DELETE FROM brands WHERE id=?", (b_id,))
                conn.commit()
                conn.close()
                QMessageBox.information(self, "Deleted", "Brand deleted successfully.")
                self.load_brands()
                self.prepare_for_new()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete brand:\n{e}")


class SRProductsArchiveDialog(QDialog):
    """Dialog for loading and viewing Excel sheets from Dropbox"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("SR Products Archive")
        self.setMinimumSize(1200, 750)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinMaxButtonsHint)
        self.sheets_data = {}
        self.raw_excel_content = None
        self.processed_sheets = {}
        self.selected_file_path = None
        self.selected_file_type = None
        self.selected_vendor_name = None
        self.cache_dir = Path(__file__).resolve().parent / "data" / "cache"
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.force_fresh_fetch = False
        self._build_ui()
        
        # Load last URL if present
        last_url_path = self.cache_dir / "last_url.txt"
        if last_url_path.exists():
            try:
                self.url_input.setText(last_url_path.read_text(encoding='utf-8').strip())
            except Exception as e:
                print(f"Error loading last URL: {e}")
                
        # Auto-load default session if present
        self.auto_load_default_session()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        
        # URL input section
        url_layout = QHBoxLayout()
        url_layout.addWidget(QLabel("Dropbox Excel Sheet URL:"))
        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("https://www.dropbox.com/...")
        url_layout.addWidget(self.url_input)
        
        load_btn = QPushButton("Load Sheet")
        load_btn.clicked.connect(self.load_sheet)
        url_layout.addWidget(load_btn)

        self.fetch_new_btn = QPushButton("Fetch New Data")
        self.fetch_new_btn.setStyleSheet("background-color: #ffc107; color: black; font-weight: bold;")
        self.fetch_new_btn.clicked.connect(self.fetch_new_data)
        url_layout.addWidget(self.fetch_new_btn)

        self.load_versioned_btn = QPushButton("Load Versioned")
        self.load_versioned_btn.setStyleSheet("background-color: #6f42c1; color: white; font-weight: bold;")
        self.load_versioned_btn.clicked.connect(self.load_versioned_session)
        url_layout.addWidget(self.load_versioned_btn)

        self.save_version_btn = QPushButton("Save Session Version")
        self.save_version_btn.setStyleSheet("background-color: #17a2b8; color: white; font-weight: bold;")
        self.save_version_btn.clicked.connect(self.save_session_version)
        self.save_version_btn.setVisible(False)
        url_layout.addWidget(self.save_version_btn)

        self.fullscreen_btn = QPushButton("Fullscreen")
        self.fullscreen_btn.setStyleSheet("background-color: #6c757d; color: white;")
        self.fullscreen_btn.clicked.connect(self.toggle_fullscreen)
        url_layout.addWidget(self.fullscreen_btn)

        self.download_btn = QPushButton("Download Processed")
        self.download_btn.setStyleSheet("background-color: #17a2b8; color: white;")
        self.download_btn.clicked.connect(self.download_processed)
        self.download_btn.setVisible(False)
        url_layout.addWidget(self.download_btn)

        self.view_tag_list_btn = QPushButton("View Tag List")
        self.view_tag_list_btn.setStyleSheet("background-color: #007bff; color: white; font-weight: bold;")
        self.view_tag_list_btn.clicked.connect(self.view_tag_list)
        self.view_tag_list_btn.setVisible(False)
        url_layout.addWidget(self.view_tag_list_btn)

        layout.addLayout(url_layout)

        # Main splitter (vertical)
        self.main_splitter = QSplitter(Qt.Vertical)
        self.main_splitter.setVisible(False)
        layout.addWidget(self.main_splitter)
        
        # Upper container for sheets and table search filter (upper half)
        self.upper_container = QWidget()
        upper_layout = QVBoxLayout(self.upper_container)
        upper_layout.setContentsMargins(0, 0, 0, 0)
        
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("<b>Table Filter Search:</b>"))
        self.tab_search_input = QLineEdit()
        self.tab_search_input.setPlaceholderText("Type to search/filter rows in the visible table...")
        self.tab_search_input.textChanged.connect(self.filter_current_tab_table)
        filter_layout.addWidget(self.tab_search_input)
        upper_layout.addLayout(filter_layout)

        # Checkable Dropdown Filters Layout
        dropdown_layout = QHBoxLayout()
        self.brand_filter_label = QLabel("<b>Brand:</b>")
        self.brand_filter = CheckableComboBox()
        self.brand_filter.currentTextChanged.connect(self.filter_current_tab_table)
        
        self.category_filter_label = QLabel("<b>Category:</b>")
        self.category_filter = CheckableComboBox()
        self.category_filter.currentTextChanged.connect(self.filter_current_tab_table)
        
        self.subcat_filter_label = QLabel("<b>Sub-Category:</b>")
        self.subcat_filter = CheckableComboBox()
        self.subcat_filter.currentTextChanged.connect(self.filter_current_tab_table)
        
        self.vendor_filter_label = QLabel("<b>Vendor:</b>")
        self.vendor_filter = CheckableComboBox()
        self.vendor_filter.currentTextChanged.connect(self.filter_current_tab_table)
        
        self.date_filter_label = QLabel("<b>Invoice Date:</b>")
        self.date_filter = CheckableComboBox()
        self.date_filter.currentTextChanged.connect(self.filter_current_tab_table)
        
        dropdown_layout.addWidget(self.brand_filter_label)
        dropdown_layout.addWidget(self.brand_filter)
        dropdown_layout.addWidget(self.category_filter_label)
        dropdown_layout.addWidget(self.category_filter)
        dropdown_layout.addWidget(self.subcat_filter_label)
        dropdown_layout.addWidget(self.subcat_filter)
        dropdown_layout.addWidget(self.vendor_filter_label)
        dropdown_layout.addWidget(self.vendor_filter)
        dropdown_layout.addWidget(self.date_filter_label)
        dropdown_layout.addWidget(self.date_filter)
        upper_layout.addLayout(dropdown_layout)
        
        # Hide filters by default
        self.brand_filter_label.setVisible(False)
        self.brand_filter.setVisible(False)
        self.category_filter_label.setVisible(False)
        self.category_filter.setVisible(False)
        self.subcat_filter_label.setVisible(False)
        self.subcat_filter.setVisible(False)
        self.vendor_filter_label.setVisible(False)
        self.vendor_filter.setVisible(False)
        self.date_filter_label.setVisible(False)
        self.date_filter.setVisible(False)
        
        self.tab_widget = QTabWidget()
        self.tab_widget.currentChanged.connect(self.on_tab_changed)
        upper_layout.addWidget(self.tab_widget)
        
        self.main_splitter.addWidget(self.upper_container)
        
        # Bottom panel (lower half)
        self.bottom_panel = QWidget()
        bottom_layout = QHBoxLayout(self.bottom_panel)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        self.main_splitter.addWidget(self.bottom_panel)

        # Splitter for bottom panel (left 1/4, right 3/4)
        self.bottom_splitter = QSplitter(Qt.Horizontal)
        bottom_layout.addWidget(self.bottom_splitter)

        # Left 1/4 widget (Vendors and files list)
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 5, 0)

        left_layout.addWidget(QLabel("<b>Available Vendors:</b>"))
        self.vendor_list = QListWidget()
        self.vendor_list.itemSelectionChanged.connect(self.on_vendor_selected)
        left_layout.addWidget(self.vendor_list)

        left_layout.addWidget(QLabel("<b>Vendor Files:</b>"))
        self.vendor_files_list = QListWidget()
        self.vendor_files_list.itemSelectionChanged.connect(self.on_file_selected)
        left_layout.addWidget(self.vendor_files_list)

        self.bottom_splitter.addWidget(left_widget)

        # Right 3/4 widget (Standardization Control Panel)
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(5, 0, 0, 0)

        right_layout.addWidget(QLabel("<b>Standardization Control Panel</b>"))
        
        info_group = QGroupBox("Selection Details")
        info_layout = QVBoxLayout(info_group)
        self.info_vendor_label = QLabel("Selected Vendor: None")
        self.info_file_label = QLabel("Selected File: None")
        self.info_type_label = QLabel("File Type: None")
        info_layout.addWidget(self.info_vendor_label)
        info_layout.addWidget(self.info_file_label)
        info_layout.addWidget(self.info_type_label)
        right_layout.addWidget(info_group)

        # Mode Selection
        mode_group = QGroupBox("Selected File Mode Option")
        mode_layout = QHBoxLayout(mode_group)
        self.radio_merge = QRadioButton("Merge File")
        self.radio_unique = QRadioButton("Unique File")
        self.radio_merge.setEnabled(False)
        self.radio_unique.setEnabled(False)
        mode_layout.addWidget(self.radio_merge)
        mode_layout.addWidget(self.radio_unique)
        right_layout.addWidget(mode_group)

        # Button layout
        right_layout.addStretch()
        self.apply_standards_btn = QPushButton("Apply Sunrise Standard")
        self.apply_standards_btn.setEnabled(False)
        self.apply_standards_btn.setStyleSheet(
            "background-color: #28a745; color: white; font-size: 13pt; font-weight: bold; padding: 12px; border-radius: 5px;"
        )
        self.apply_standards_btn.clicked.connect(self.apply_sunrise_standard)
        right_layout.addWidget(self.apply_standards_btn)

        self.bottom_splitter.addWidget(right_widget)

        # Set stretch factors for bottom splitter: left 1, right 3 (1/4 and 3/4)
        self.bottom_splitter.setStretchFactor(0, 1)
        self.bottom_splitter.setStretchFactor(1, 3)
        
        # Status/message label
        self.status_label = QLabel("Enter a Dropbox URL and click 'Load Sheet'")
        layout.addWidget(self.status_label)

    def toggle_fullscreen(self) -> None:
        if self.isFullScreen():
            self.showNormal()
            self.fullscreen_btn.setText("Fullscreen")
        else:
            self.showFullScreen()
            self.fullscreen_btn.setText("Normal Screen")

    def filter_current_tab_table(self) -> None:
        query = self.tab_search_input.text().strip().lower()
        current_table = self.tab_widget.currentWidget()
        if not isinstance(current_table, QTableWidget):
            return

        # Get active filters
        selected_brands = self.brand_filter.checked_items() if hasattr(self, 'brand_filter') else []
        selected_categories = self.category_filter.checked_items() if hasattr(self, 'category_filter') else []
        selected_subcats = self.subcat_filter.checked_items() if hasattr(self, 'subcat_filter') else []
        selected_vendors = self.vendor_filter.checked_items() if hasattr(self, 'vendor_filter') else []
        selected_dates = self.date_filter.checked_items() if hasattr(self, 'date_filter') else []

        # Get column indices
        brand_col = -1
        cat_col = -1
        subcat_col = -1
        vendor_col = -1
        date_col = -1
        for col_idx in range(current_table.columnCount()):
            header = current_table.horizontalHeaderItem(col_idx)
            if header:
                header_text = header.text().strip().lower().replace(' ', '').replace('_', '').replace('-', '')
                if header_text == 'brand':
                    brand_col = col_idx
                elif header_text == 'category':
                    cat_col = col_idx
                elif header_text in ['subcategory', 'subcat']:
                    subcat_col = col_idx
                elif header_text in ['vendor', 'vendorname']:
                    vendor_col = col_idx
                elif header_text in ['invoicedate', 'date', 'datum']:
                    date_col = col_idx

        # Hide or show filters depending on whether the columns exist in current table
        if hasattr(self, 'brand_filter'):
            has_brand = (brand_col != -1)
            self.brand_filter.setVisible(has_brand)
            self.brand_filter_label.setVisible(has_brand)

            has_cat = (cat_col != -1)
            self.category_filter.setVisible(has_cat)
            self.category_filter_label.setVisible(has_cat)

            has_subcat = (subcat_col != -1)
            self.subcat_filter.setVisible(has_subcat)
            self.subcat_filter_label.setVisible(has_subcat)

            has_vendor = (vendor_col != -1)
            self.vendor_filter.setVisible(has_vendor)
            self.vendor_filter_label.setVisible(has_vendor)

            has_date = (date_col != -1)
            self.date_filter.setVisible(has_date)
            self.date_filter_label.setVisible(has_date)

        for row in range(current_table.rowCount()):
            # 1. Text Search Filter
            text_match = False
            if not query:
                text_match = True
            else:
                for col in range(current_table.columnCount()):
                    header = current_table.horizontalHeaderItem(col)
                    if header and header.text() == 'In Tag Bucket':
                        continue
                    item = current_table.item(row, col)
                    if item and query in item.text().lower():
                        text_match = True
                        break

            # 2. Brand Filter
            brand_match = True
            if selected_brands and brand_col != -1:
                item = current_table.item(row, brand_col)
                val = item.text().strip() if item else ""
                brand_match = (val in selected_brands)

            # 3. Category Filter
            cat_match = True
            if selected_categories and cat_col != -1:
                item = current_table.item(row, cat_col)
                val = item.text().strip() if item else ""
                cat_match = (val in selected_categories)

            # 4. Sub-Category Filter
            subcat_match = True
            if selected_subcats and subcat_col != -1:
                item = current_table.item(row, subcat_col)
                val = item.text().strip() if item else ""
                subcat_match = (val in selected_subcats)

            # 5. Vendor Filter
            vendor_match = True
            if selected_vendors and vendor_col != -1:
                item = current_table.item(row, vendor_col)
                val = item.text().strip() if item else ""
                vendor_match = (val in selected_vendors)

            # 6. Invoice Date Filter
            date_match = True
            if selected_dates and date_col != -1:
                item = current_table.item(row, date_col)
                val = item.text().strip() if item else ""
                date_match = (val in selected_dates)

            # Show row only if all conditions match
            current_table.setRowHidden(row, not (text_match and brand_match and cat_match and subcat_match and vendor_match and date_match))

    def on_tab_changed(self, index) -> None:
        self.populate_dropdown_filters()
        self.filter_current_tab_table()

    def populate_dropdown_filters(self) -> None:
        current_table = self.tab_widget.currentWidget()
        if not isinstance(current_table, QTableWidget):
            return

        self.brand_filter.blockSignals(True)
        self.category_filter.blockSignals(True)
        self.subcat_filter.blockSignals(True)
        self.vendor_filter.blockSignals(True)
        self.date_filter.blockSignals(True)

        brand_col = -1
        cat_col = -1
        subcat_col = -1
        vendor_col = -1
        date_col = -1
        for col_idx in range(current_table.columnCount()):
            header = current_table.horizontalHeaderItem(col_idx)
            if header:
                header_text = header.text().strip().lower().replace(' ', '').replace('_', '').replace('-', '')
                if header_text == 'brand':
                    brand_col = col_idx
                elif header_text == 'category':
                    cat_col = col_idx
                elif header_text in ['subcategory', 'subcat']:
                    subcat_col = col_idx
                elif header_text in ['vendor', 'vendorname']:
                    vendor_col = col_idx
                elif header_text in ['invoicedate', 'date', 'datum']:
                    date_col = col_idx

        brands = set()
        categories = set()
        subcats = set()
        vendors = set()
        dates = set()

        for row in range(current_table.rowCount()):
            if brand_col != -1:
                item = current_table.item(row, brand_col)
                if item and item.text().strip():
                    brands.add(item.text().strip())
            if cat_col != -1:
                item = current_table.item(row, cat_col)
                if item and item.text().strip():
                    categories.add(item.text().strip())
            if subcat_col != -1:
                item = current_table.item(row, subcat_col)
                if item and item.text().strip():
                    subcats.add(item.text().strip())
            if vendor_col != -1:
                item = current_table.item(row, vendor_col)
                if item and item.text().strip():
                    vendors.add(item.text().strip())
            if date_col != -1:
                item = current_table.item(row, date_col)
                if item and item.text().strip():
                    dates.add(item.text().strip())

        if brand_col != -1:
            self.brand_filter.add_items(sorted(list(brands)))
        else:
            self.brand_filter.clear()

        if cat_col != -1:
            self.category_filter.add_items(sorted(list(categories)))
        else:
            self.category_filter.clear()

        if subcat_col != -1:
            self.subcat_filter.add_items(sorted(list(subcats)))
        else:
            self.subcat_filter.clear()

        if vendor_col != -1:
            self.vendor_filter.add_items(sorted(list(vendors)))
        else:
            self.vendor_filter.clear()

        if date_col != -1:
            def parse_date_key(d_str):
                if not d_str or str(d_str).strip() in ("", "nan", "None", "N/A"):
                    return pd.Timestamp.min
                try:
                    dt = pd.to_datetime(d_str, dayfirst=True, errors='coerce')
                    if pd.isnull(dt):
                        return pd.Timestamp.min
                    return dt
                except Exception:
                    return pd.Timestamp.min
            sorted_dates = sorted(list(dates), key=parse_date_key, reverse=True)
            self.date_filter.add_items(sorted_dates)
        else:
            self.date_filter.clear()

        self.brand_filter.blockSignals(False)
        self.category_filter.blockSignals(False)
        self.subcat_filter.blockSignals(False)
        self.vendor_filter.blockSignals(False)
        self.date_filter.blockSignals(False)

    def view_tag_list(self) -> None:
        current_table = self.tab_widget.currentWidget()
        if not isinstance(current_table, QTableWidget):
            QMessageBox.warning(self, "Error", "No active sheet table found.")
            return

        chk_col = -1
        for col in range(current_table.columnCount()):
            header = current_table.horizontalHeaderItem(col)
            if header and header.text() == 'In Tag Bucket':
                chk_col = col
                break

        if chk_col == -1:
            QMessageBox.warning(self, "Error", "The active sheet does not support row selection ('In Tag Bucket' column not found).")
            return

        selected_rows = []
        for row in range(current_table.rowCount()):
            item = current_table.item(row, chk_col)
            if item and item.checkState() == Qt.Checked:
                row_data = {}
                for col in range(current_table.columnCount()):
                    if col == chk_col:
                        continue
                    h_item = current_table.horizontalHeaderItem(col)
                    header_name = h_item.text() if h_item else f"Col_{col}"
                    cell_item = current_table.item(row, col)
                    row_data[header_name] = cell_item.text() if cell_item else ""
                selected_rows.append(row_data)

        if not selected_rows:
            QMessageBox.warning(self, "No Selection", "Please select at least one product in the 'In Tag Bucket' column.")
            return

        dialog = TagListDialog(selected_rows, self)
        dialog.exec_()

    def auto_load_default_session(self) -> None:
        default_file = self.cache_dir / "default_session.xlsx"
        if default_file.exists():
            try:
                self.status_label.setText("Auto-loading default session...")
                with open(default_file, "rb") as f:
                    excel_bytes = f.read()
                
                self.raw_excel_content = excel_bytes
                excel_file = io.BytesIO(excel_bytes)
                xls = pd.ExcelFile(excel_file, engine='openpyxl')
                
                self.sheets_data = {}
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
                    self.sheets_data[sheet_name] = df
                
                self._display_sheets()
                self.load_vendors()
                self.status_label.setText("Successfully loaded default session.")
                
                self.main_splitter.setVisible(True)
                self.main_splitter.setSizes([350, 350])
                self.download_btn.setVisible(False)
                self.view_tag_list_btn.setVisible(False)
                self.save_version_btn.setVisible(True)
            except Exception as e:
                self.status_label.setText(f"Error auto-loading default session: {e}")

    def load_vendors(self) -> None:
        self.vendor_list.clear()
        if VENDOR_ROOT.exists():
            vendor_names = [p.name for p in VENDOR_ROOT.iterdir() if p.is_dir()]
            vendor_names.sort()
            for vendor in vendor_names:
                self.vendor_list.addItem(vendor)

    def on_vendor_selected(self) -> None:
        selected_items = self.vendor_list.selectedItems()
        if not selected_items:
            self.vendor_files_list.clear()
            self.info_vendor_label.setText("Selected Vendor: None")
            return
            
        vendor_name = selected_items[0].text()
        self.selected_vendor_name = vendor_name
        self.info_vendor_label.setText(f"Selected Vendor: {vendor_name}")
        
        self.vendor_files_list.clear()
        
        # Determine vendor specific files
        vendor_lower = vendor_name.lower()
        workspace_root = Path(r"c:\Users\mdtou\PycharmProjects\sr-vendor-handlers")
        
        merged_path = workspace_root / f"merged_{vendor_lower}.xlsx"
        if not merged_path.exists():
            merged_path = workspace_root / "merged.xlsx"
            
        unique_path = workspace_root / f"unique_{vendor_lower}.xlsx"
        if not unique_path.exists():
            unique_path = workspace_root / f"{vendor_lower}_unique_products.xlsx"
            
        unique_exists = unique_path.exists()
        if not unique_exists:
            # Fallback: auto calculated from merged file
            unique_path = merged_path
            
        # Add items to the files list if they exist, or display with a warning
        merged_item = QListWidgetItem(f"Merged Purchased Archives ({merged_path.name})")
        merged_item.setData(Qt.UserRole, str(merged_path))
        merged_item.setData(Qt.UserRole + 1, "merge")
        if not merged_path.exists():
            merged_item.setText(f"Merged Purchased Archives (Not Found)")
            merged_item.setFlags(merged_item.flags() & ~Qt.ItemIsEnabled)
        self.vendor_files_list.addItem(merged_item)
        
        unique_item_name = f"Unique Products ({unique_path.name})" if unique_exists else f"Unique Products (Auto-calculate from Merged)"
        unique_item = QListWidgetItem(unique_item_name)
        unique_item.setData(Qt.UserRole, str(unique_path))
        unique_item.setData(Qt.UserRole + 1, "unique")
        # Always enable the unique option if unique_exists or merged_path exists
        if not unique_exists and not merged_path.exists():
            unique_item.setText(f"Unique Products (Not Found)")
            unique_item.setFlags(unique_item.flags() & ~Qt.ItemIsEnabled)
        self.vendor_files_list.addItem(unique_item)

    def on_file_selected(self) -> None:
        selected_items = self.vendor_files_list.selectedItems()
        if not selected_items:
            self.apply_standards_btn.setEnabled(False)
            self.info_file_label.setText("Selected File: None")
            self.info_type_label.setText("File Type: None")
            self.radio_merge.setChecked(False)
            self.radio_unique.setChecked(False)
            return
            
        item = selected_items[0]
        file_path = item.data(Qt.UserRole)
        file_type = item.data(Qt.UserRole + 1)
        
        self.selected_file_path = file_path
        self.selected_file_type = file_type
        
        self.info_file_label.setText(f"Selected File: {Path(file_path).name}")
        self.info_type_label.setText(f"File Type: {file_type.capitalize()}")
        
        # Check radio button
        self.radio_merge.setEnabled(True)
        self.radio_unique.setEnabled(True)
        if file_type == "merge":
            self.radio_merge.setChecked(True)
        else:
            self.radio_unique.setChecked(True)
            
        self.apply_standards_btn.setEnabled(True)

    def fetch_new_data(self) -> None:
        self.force_fresh_fetch = True
        self.load_sheet()

    def save_session_version(self) -> None:
        if not self.raw_excel_content:
            QMessageBox.warning(self, "Error", "No loaded Excel content to save.")
            return
            
        text, ok = QInputDialog.getText(self, "Save Loaded Session Version", "Enter version name for this loaded spreadsheet:")
        if ok and text:
            version_name = text.strip()
            # Clean version name to be a safe filename
            import re
            version_name_clean = re.sub(r'[^a-zA-Z0-9_\-]', '_', version_name)
            if not version_name_clean:
                QMessageBox.warning(self, "Error", "Invalid version name.")
                return
                
            versions_dir = self.cache_dir / "versions"
            versions_dir.mkdir(parents=True, exist_ok=True)
            version_file_path = versions_dir / f"{version_name_clean}.xlsx"
            
            # If exists, ask for overwrite confirmation
            if version_file_path.exists():
                reply = QMessageBox.question(
                    self, 
                    "Overwrite Version", 
                    f"A version named '{version_name_clean}' already exists. Overwrite it?",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.No:
                    return
                    
            try:
                with open(version_file_path, "wb") as f:
                    f.write(self.raw_excel_content)
                # Also save to default_session.xlsx
                try:
                    with open(self.cache_dir / "default_session.xlsx", "wb") as f:
                        f.write(self.raw_excel_content)
                except Exception as ex:
                    print(f"Error updating default_session.xlsx: {ex}")
                QMessageBox.information(self, "Success", f"Session saved successfully as version '{version_name_clean}'.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save version:\n{str(e)}")

    def load_versioned_session(self) -> None:
        versions_dir = self.cache_dir / "versions"
        if not versions_dir.exists() or not list(versions_dir.glob("*.xlsx")):
            QMessageBox.information(self, "No Versions Found", "No saved session versions found on disk.")
            return
            
        # Get list of versions
        version_files = sorted(list(versions_dir.glob("*.xlsx")))
        versions_list = [f.stem for f in version_files]
        
        item, ok = QInputDialog.getItem(self, "Load Versioned Session", "Select a saved version to load:", versions_list, 0, False)
        if ok and item:
            selected_file = versions_dir / f"{item}.xlsx"
            try:
                self.status_label.setText(f"Loading version '{item}'...")
                self.main_splitter.setVisible(False)
                
                with open(selected_file, "rb") as f:
                    excel_bytes = f.read()
                    
                self.raw_excel_content = excel_bytes
                
                # Parse Excel
                excel_file = io.BytesIO(excel_bytes)
                xls = pd.ExcelFile(excel_file, engine='openpyxl')
                
                self.sheets_data = {}
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
                    self.sheets_data[sheet_name] = df
                
                self._display_sheets()
                self.load_vendors()
                
                # Save as default session
                try:
                    with open(self.cache_dir / "default_session.xlsx", "wb") as f:
                        f.write(excel_bytes)
                    last_url_path = self.cache_dir / "last_url.txt"
                    if last_url_path.exists():
                        last_url_path.unlink()
                except Exception as ex:
                    print(f"Error saving default session: {ex}")
                
                self.status_label.setText(f"Successfully loaded version '{item}' from local storage")
                
                # Show layout
                self.main_splitter.setVisible(True)
                self.main_splitter.setSizes([350, 350])
                self.download_btn.setVisible(False)
                self.view_tag_list_btn.setVisible(False)
                self.save_version_btn.setVisible(True)
                
            except Exception as e:
                self.status_label.setText(f"Error loading version: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to load version '{item}':\n{str(e)}")

    def load_sheet(self) -> None:
        url = self.url_input.text().strip()
        if not url:
            QMessageBox.warning(self, "Error", "Please enter a Dropbox URL.")
            return
        
        self.status_label.setText("Loading...")
        self.main_splitter.setVisible(False)
        
        # Determine cache file path for this URL
        import hashlib
        url_hash = hashlib.sha256(url.encode('utf-8')).hexdigest()
        cache_file_path = self.cache_dir / f"{url_hash}.xlsx"
        
        force_fetch = getattr(self, "force_fresh_fetch", False)
        self.force_fresh_fetch = False # reset
        
        excel_bytes = None
        loaded_from_cache = False
        
        if cache_file_path.exists() and not force_fetch:
            try:
                with open(cache_file_path, "rb") as f:
                    excel_bytes = f.read()
                loaded_from_cache = True
            except Exception as e:
                print(f"Error reading cache file: {e}")
                
        if excel_bytes is None:
            try:
                # Convert Dropbox sharing URL to direct download
                if 'dropbox.com' in url:
                    url = url.replace('dl=0', 'dl=1')
                    if 'dl=1' not in url:
                        if '?' in url:
                            url += '&dl=1'
                        else:
                            url += '?dl=1'
                
                # Fetch file with cache-busting logic if force_fetch is requested
                headers = {'User-Agent': 'Mozilla/5.0'}
                fetch_url = url
                if force_fetch:
                    headers['Cache-Control'] = 'no-cache'
                    headers['Pragma'] = 'no-cache'
                    import time
                    cb = int(time.time())
                    if '?' in fetch_url:
                        fetch_url += f"&cb={cb}"
                    else:
                        fetch_url += f"?cb={cb}"
                
                response = requests.get(fetch_url, headers=headers, timeout=30)
                response.raise_for_status()
                excel_bytes = response.content
                
                # Save to cache
                self.cache_dir.mkdir(parents=True, exist_ok=True)
                with open(cache_file_path, "wb") as f:
                    f.write(excel_bytes)
                
            except requests.exceptions.RequestException as e:
                self.status_label.setText(f"Error: Failed to fetch file - {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to fetch file:\n{str(e)}")
                return
            except Exception as e:
                self.status_label.setText(f"Error: Failed to write cache - {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to write cache:\n{str(e)}")
                return
        
        try:
            self.raw_excel_content = excel_bytes
            
            # Parse Excel
            excel_file = io.BytesIO(excel_bytes)
            xls = pd.ExcelFile(excel_file, engine='openpyxl')
            
            self.sheets_data = {}
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
                self.sheets_data[sheet_name] = df
            
            self._display_sheets()
            self.load_vendors()
            
            # Save as default session and save URL
            try:
                with open(self.cache_dir / "default_session.xlsx", "wb") as f:
                    f.write(excel_bytes)
                with open(self.cache_dir / "last_url.txt", "w", encoding='utf-8') as f:
                    f.write(url)
            except Exception as ex:
                print(f"Error saving default session or URL: {ex}")
            
            if loaded_from_cache:
                self.status_label.setText(f"Successfully loaded {len(self.sheets_data)} sheet(s) from Cache (Materialized View)")
            else:
                self.status_label.setText(f"Successfully fetched and cached {len(self.sheets_data)} sheet(s)")
            
            # Show the splitter and set sizing (50% / 50% split)
            self.main_splitter.setVisible(True)
            self.main_splitter.setSizes([350, 350])
            self.download_btn.setVisible(False)
            self.view_tag_list_btn.setVisible(False)
            self.save_version_btn.setVisible(True)
            
        except Exception as e:
            self.status_label.setText(f"Error: Failed to parse file - {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to parse Excel file:\n{str(e)}")

    def _display_sheets(self, data_dict=None) -> None:
        self.tab_widget.clear()
        
        display_data = data_dict if data_dict is not None else self.sheets_data
        for sheet_name, df in display_data.items():
            # Create table for this sheet
            table = QTableWidget()
            is_unique_products = (sheet_name == 'unique products')
            extra_cols = 1 if is_unique_products else 0
            
            table.setColumnCount(len(df.columns) + extra_cols)
            table.setRowCount(len(df))
            
            headers = []
            if is_unique_products:
                headers.append('In Tag Bucket')
            headers.extend([str(col) for col in df.columns])
            table.setHorizontalHeaderLabels(headers)
            
            # Use interactive resize mode so it is readable if many columns
            if (len(df.columns) + extra_cols) > 8:
                table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
            else:
                table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            
            # Fill table with data
            for row_idx, (_, row_data) in enumerate(df.iterrows()):
                if is_unique_products:
                    chk_item = QTableWidgetItem()
                    chk_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    chk_item.setCheckState(Qt.Unchecked)
                    table.setItem(row_idx, 0, chk_item)
                    
                for col_idx, value in enumerate(row_data):
                    val_str = str(value) if value is not None and str(value) != 'nan' else ''
                    item = QTableWidgetItem(val_str)
                    if val_str:
                        item.setToolTip(val_str)
                    table.setItem(row_idx, col_idx + extra_cols, item)
            
            if is_unique_products:
                table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
                table.setColumnWidth(0, 100)
                
            # Set specific column widths for Item, Name, and Description
            for col_idx, col in enumerate(df.columns):
                c_name = str(col).lower().replace(' ', '').replace('_', '')
                if c_name in ['item', 'name', 'description']:
                    table.horizontalHeader().setSectionResizeMode(col_idx + extra_cols, QHeaderView.Interactive)
                    table.setColumnWidth(col_idx + extra_cols, 350)
            
            # Add table to tab widget
            self.tab_widget.addTab(table, sheet_name)
            
        self.on_tab_changed(self.tab_widget.currentIndex())

    def apply_sunrise_standard(self) -> None:
        if not self.raw_excel_content:
            QMessageBox.warning(self, "Error", "No loaded Excel content.")
            return
            
        if not self.selected_file_path or not Path(self.selected_file_path).exists():
            QMessageBox.warning(self, "Error", "Please select a valid vendor file first.")
            return
            
        if 'All' not in self.sheets_data:
            QMessageBox.warning(self, "Error", "The loaded Excel does not contain a sheet named 'All'.")
            return
            
        self.status_label.setText("Applying standards...")
        self.apply_standards_btn.setEnabled(False)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        
        try:
            # Helper functions for robust column resolution
            def get_row_item_no(r):
                for k in ['Item No.', 'Art No', 'Item No', 'ArtNo', 'article_no', 'sku']:
                    val = r.get(k)
                    if pd.notna(val) and str(val).strip() != '' and str(val).strip().lower() != 'nan':
                        val_str = str(val).strip()
                        if val_str.endswith('.0'):
                            val_str = val_str[:-2]
                        return val_str
                return ''

            def get_row_desc(r):
                for k in ['Description', 'item', 'Name', 'product_name', 'item_name']:
                    val = r.get(k)
                    if pd.notna(val) and str(val).strip() != '' and str(val).strip().lower() != 'nan':
                        return str(val).strip()
                return ''

            def get_matched_unit_price(mr):
                for k in ['/pc preis', '/pc pres', 'unit_price', '/pcpreis']:
                    if k in mr:
                        return mr[k]
                for k in mr.index:
                    if '/pc' in str(k).lower() or 'preis' in str(k).lower():
                        return mr[k]
                return None

            def get_matched_sale_price(mr):
                for k in ['Sale price ', 'sale_price', 'Sale price', 'sale price']:
                    if k in mr:
                        return mr[k]
                cols = list(mr.index)
                for k in ['/pc preis', '/pc pres', '/pcpreis']:
                    if k in cols:
                        idx = cols.index(k)
                        if idx + 1 < len(cols):
                            return mr.iloc[idx + 1]
                return None

            def get_matched_margin_50(mr):
                for k in ['margin_50', 'margin', 'Margin']:
                    if k in mr:
                        return mr[k]
                cols = list(mr.index)
                for k in ['/pc preis', '/pc pres', '/pcpreis']:
                    if k in cols:
                        idx = cols.index(k)
                        if idx + 2 < len(cols):
                            return mr.iloc[idx + 2]
                return None

            # Determine correct source sheet based on selection
            df_source = None
            product_sheet_name = 'All'
            if self.selected_file_type == 'unique':
                # Load the product sheet of the loaded excel (sheet next to All)
                sheet_names = list(self.sheets_data.keys())
                product_sheet_name = None
                if 'All' in sheet_names:
                    all_idx = sheet_names.index('All')
                    if all_idx + 1 < len(sheet_names):
                        product_sheet_name = sheet_names[all_idx + 1]
                if not product_sheet_name:
                    for s in sheet_names:
                        if 'product' in s.lower():
                            product_sheet_name = s
                            break
                if not product_sheet_name and len(sheet_names) > 1:
                    product_sheet_name = sheet_names[1]
                if not product_sheet_name:
                    product_sheet_name = 'All'
                df_source = self.sheets_data[product_sheet_name]
            else:
                df_source = self.sheets_data['All']

            # Load the selected purchase archives merge sheet (original)
            selected_path = Path(self.selected_file_path)
            with pd.ExcelFile(selected_path) as xls:
                # Determine appropriate sheet name (robust lookup)
                sheet_name = None
                pref_sheets = ['unique products', 'sr standard archives', 'invoice data', 'purchase archives', 'all', 'produkt', 'product']
                for pref in pref_sheets:
                    for s in xls.sheet_names:
                        if s.lower() == pref:
                            sheet_name = s
                            break
                    if sheet_name:
                        break
                if not sheet_name:
                    sheet_name = xls.sheet_names[0]
                df_original_merge = pd.read_excel(xls, sheet_name=sheet_name)
            
            df_merge_processed = df_original_merge.copy()

            # 1. Rearrange from oldest order (sort chronologically)
            date_col_name = None
            for col in df_merge_processed.columns:
                if str(col).lower().replace(' ', '').replace('_', '').replace('.', '') in ['invoicedate', 'datum', 'date']:
                    date_col_name = col
                    break
                    
            if date_col_name:
                # Parse temporary datetime series for sorting
                df_merge_processed['_temp_parsed_date'] = pd.to_datetime(
                    df_merge_processed[date_col_name],
                    dayfirst=True,
                    errors='coerce'
                )
                # Sort ascending, keeping original index order for ties (stable sort)
                df_merge_processed.sort_values(by='_temp_parsed_date', ascending=True, kind='stable', inplace=True)
                df_merge_processed.reset_index(drop=True, inplace=True)
                df_merge_processed.drop(columns=['_temp_parsed_date'], inplace=True)

            # 2. Insert 'Vendor' column value in between column 'Invoice Date' and 'Invoice Number'
            if 'Vendor' in df_merge_processed.columns:
                df_merge_processed.drop(columns=['Vendor'], inplace=True)
                
            date_idx = None
            num_idx = None
            for i, col in enumerate(df_merge_processed.columns):
                c_low = str(col).lower().replace(' ', '').replace('_', '').replace('.', '')
                if c_low in ['invoicedate', 'datum', 'date']:
                    date_idx = i
                elif c_low in ['invoicenumber', 'rechnnr', 'docnr', 'invoiceno']:
                    num_idx = i
                    
            if date_idx is not None:
                insert_idx = date_idx + 1
            elif num_idx is not None:
                insert_idx = num_idx
            else:
                insert_idx = min(2, len(df_merge_processed.columns))
                
            df_merge_processed.insert(insert_idx, 'Vendor', None)
            
            # Convert all columns to object type to prevent strict Arrow/pandas 2.x dtype Checks
            for col in df_merge_processed.columns:
                df_merge_processed[col] = df_merge_processed[col].astype(object)

            # 3. Populate Vendor column values based on chronological invoice order
            vendor_lower = self.selected_vendor_name.lower() if self.selected_vendor_name else "vendor"
            base_name = "asian" if "asia" in vendor_lower else vendor_lower
            
            invoice_to_vendor_val = {}
            unique_invoice_counter = 0
            
            for idx in range(len(df_merge_processed)):
                row_merge = df_merge_processed.iloc[idx]
                inv_num = str(row_merge.get('Invoice Number', '')).strip()
                if not inv_num or inv_num.lower() == 'nan':
                    # Fallback check
                     for col in row_merge.index:
                         if str(col).lower().replace(' ', '').replace('_', '') in ['invoicenumber', 'rechnnr', 'docnr', 'invoiceno']:
                             inv_num = str(row_merge[col]).strip()
                             break
                if not inv_num or inv_num.lower() == 'nan':
                    inv_num = f"UNKNOWN_{idx}"
                    
                if inv_num not in invoice_to_vendor_val:
                    if unique_invoice_counter == 0:
                        invoice_to_vendor_val[inv_num] = base_name
                    else:
                        invoice_to_vendor_val[inv_num] = f"{base_name}{unique_invoice_counter}"
                    unique_invoice_counter += 1
                    
                df_merge_processed.at[idx, 'Vendor'] = invoice_to_vendor_val[inv_num]

            # 4. For unique file, keep only the first (oldest) entry of each unique product
            # so the first time entry of vendor value is considered in Vendor Column value.
            if self.selected_file_type == 'unique':
                seen_keys = set()
                rows_to_keep = []
                for idx in range(len(df_merge_processed)):
                    row = df_merge_processed.iloc[idx]
                    item_no = get_row_item_no(row).lower()
                    desc = get_row_desc(row).lower()
                    prod_key = (item_no, desc)
                    if prod_key not in seen_keys:
                        seen_keys.add(prod_key)
                        rows_to_keep.append(idx)
                df_merge_processed = df_merge_processed.iloc[rows_to_keep].reset_index(drop=True)

            if self.selected_file_type == 'unique':
                cols_to_add = ['Brand', 'Category', 'Sub-Category', 'Steur', 'unit_price', 'sale_price', 'margin_50', '7 days', 'Barcode', 'sku', 'inv', 'total cost']
            else:
                cols_to_add = ['Category', 'Sub-Category', 'Steur', 'unit_price', 'sale_price', 'margin_50', '7 days', 'Barcode']
            
            # Drop them first if they already exist
            for col in cols_to_add:
                if col in df_merge_processed.columns:
                    df_merge_processed.drop(columns=[col], inplace=True)

            # Identify columns to add after Vat%
            vat_col_name = None
            for col in df_merge_processed.columns:
                if str(col).lower().replace(' ', '') == 'vat%':
                    vat_col_name = col
                    break
                    
            if vat_col_name:
                vat_idx = df_merge_processed.columns.get_loc(vat_col_name)
            else:
                vat_idx = len(df_merge_processed.columns) - 1
                
            # Insert columns after Vat%
            current_idx = vat_idx + 1
            for col in cols_to_add:
                df_merge_processed.insert(current_idx, col, None)
                current_idx += 1
                
            # If unique file, add 3 new columns after Steur: Tag, Kassen, Rack
            if self.selected_file_type == 'unique':
                cols_to_add_unique = ['Tag', 'Kassen', 'Rack']
                for col in cols_to_add_unique:
                    if col in df_merge_processed.columns:
                        df_merge_processed.drop(columns=[col], inplace=True)
                        
                steur_idx = None
                for i, col in enumerate(df_merge_processed.columns):
                    if str(col).lower() == 'steur':
                        steur_idx = i
                        break
                if steur_idx is None:
                    steur_idx = len(df_merge_processed.columns) - 1
                    
                current_idx = steur_idx + 1
                for col in cols_to_add_unique:
                    df_merge_processed.insert(current_idx, col, None)
                    current_idx += 1

            # If Name column is not in df_merge_processed, insert it
            if 'Name' not in df_merge_processed.columns:
                desc_idx = None
                for c in ['description', 'item_name', 'item name', 'item']:
                    for col in df_merge_processed.columns:
                        if c in str(col).lower():
                            desc_idx = df_merge_processed.columns.get_loc(col)
                            break
                    if desc_idx is not None:
                        break
                if desc_idx is not None:
                    df_merge_processed.insert(desc_idx + 1, 'Name', None)
                else:
                    df_merge_processed.insert(0, 'Name', None)

            # Convert all columns to object type to prevent strict Arrow/pandas 2.x dtype Checks
            for col in df_merge_processed.columns:
                df_merge_processed[col] = df_merge_processed[col].astype(object)

            # Build dictionaries for fast matching on df_source (checking Vendor + Item/ArtNo)
            all_by_vendor_art = {}
            all_by_vendor_item = {}
            all_by_art_no = {}
            all_by_item = {}
            for _, row in df_source.iterrows():
                art_no = str(row.get('Art No', '')).strip().lower()
                if art_no.endswith('.0'):
                    art_no = art_no[:-2]
                item_desc = str(row.get('item', '')).strip().lower()
                vendor_val = str(row.get('Vendor', '')).strip().lower()
                
                # General fallback maps
                if art_no and art_no != 'nan':
                    all_by_art_no[art_no] = row
                if item_desc and item_desc != 'nan':
                    all_by_item[item_desc] = row
                    
                # Vendor specific maps
                if vendor_val and vendor_val != 'nan':
                    if art_no and art_no != 'nan':
                        all_by_vendor_art[(vendor_val, art_no)] = row
                    if item_desc and item_desc != 'nan':
                        all_by_vendor_item[(vendor_val, item_desc)] = row
                    
            # Load brands list from SQLite database
            brands_list = []
            try:
                init_brand_db()
                conn = sqlite3.connect("sr_retail_brands.db")
                cursor = conn.cursor()
                cursor.execute("SELECT id, name, logo_path, origin FROM brands ORDER BY LENGTH(name) DESC, name ASC")
                brands_list = cursor.fetchall()
                conn.close()
            except Exception as e:
                print(f"Error loading brands from DB: {e}")

            # Cache for Barcode: {artno: {item: Barcode}}
            barcode_cache = {}
            
            # First pass: compare and find matches, populate cache, write values
            for idx in range(len(df_merge_processed)):
                row_merge = df_merge_processed.iloc[idx]
                
                # Get item no, description, and vendor from merge sheet
                item_no_val = get_row_item_no(row_merge).lower()
                desc_val = get_row_desc(row_merge).lower()
                vendor_val = str(row_merge.get('Vendor', '')).strip().lower()
                
                matched_row = None
                
                # Try vendor-specific matching if vendor_val is populated
                if vendor_val and vendor_val != 'nan':
                    if item_no_val and item_no_val != 'nan' and (vendor_val, item_no_val) in all_by_vendor_art:
                        matched_row = all_by_vendor_art[(vendor_val, item_no_val)]
                    elif desc_val and desc_val != 'nan' and (vendor_val, desc_val) in all_by_vendor_item:
                        matched_row = all_by_vendor_item[(vendor_val, desc_val)]
                        
                # Fallback to general matching if no vendor match found (or vendor was not specified)
                if matched_row is None:
                    if item_no_val and item_no_val != 'nan' and item_no_val in all_by_art_no:
                        matched_row = all_by_art_no[item_no_val]
                    elif desc_val and desc_val != 'nan' and desc_val in all_by_item:
                        matched_row = all_by_item[desc_val]
                    
                if matched_row is not None:
                    # 1. Process description to update Name column value
                    name_val = matched_row.get('Name')
                    item_desc = matched_row.get('item', '')
                    if pd.isna(item_desc) or not str(item_desc).strip():
                        item_desc = get_row_desc(row_merge)
                    
                    # Check if there is a package multiplier pattern
                    has_multiplier = False
                    if item_desc:
                        has_multiplier = bool(re.search(r'\b\d+\s*[xX*]\s*(\d+(?:\.\d+)?)\s*(kg|g|l|ml|pcs|piece|pieces|oz|lbs)\b', str(item_desc), flags=re.IGNORECASE))
                    
                    # Normalize description if multiplier present
                    normalized_name = ""
                    if has_multiplier and item_desc:
                        normalized_name = normalize_description_to_name(str(item_desc))
                    
                    # 2. Extract brand using database
                    detected_brand = ""
                    if item_desc:
                        detected_brand = extract_brand_from_desc(str(item_desc), brands_list)
                    
                    target_name = normalized_name if normalized_name else name_val
                    if pd.isna(target_name):
                        target_name = ""
                    else:
                        target_name = str(target_name).strip()

                    if detected_brand and target_name:
                        start_boundary = r'\b' if detected_brand[0].isalnum() else ''
                        end_boundary = r'\b' if detected_brand[-1].isalnum() else ''
                        pattern = start_boundary + re.escape(detected_brand) + end_boundary
                        cleaned_name = re.sub(pattern, '', target_name, flags=re.IGNORECASE)
                        cleaned_name = re.sub(r'\s+', ' ', cleaned_name).strip()
                        target_name = f"{detected_brand} {cleaned_name}"
                    elif detected_brand and not target_name:
                        target_name = detected_brand

                    df_merge_processed.at[idx, 'Name'] = target_name
                    
                    # Copy standard columns
                    for col in ['Category', 'Sub-Category', '7 days']:
                        df_merge_processed.at[idx, col] = matched_row.get(col)
                        
                    # Copy Steur/Stuer
                    steur_val = matched_row.get('Steur') if 'Steur' in matched_row else matched_row.get('Stuer')
                    df_merge_processed.at[idx, 'Steur'] = steur_val
                    
                    # Copy unit_price, sale_price, margin_50 using our robust helpers
                    df_merge_processed.at[idx, 'unit_price'] = get_matched_unit_price(matched_row)
                    df_merge_processed.at[idx, 'sale_price'] = get_matched_sale_price(matched_row)
                    df_merge_processed.at[idx, 'margin_50'] = get_matched_margin_50(matched_row)
                    
                    # Copy unique columns if selected file is unique
                    if self.selected_file_type == 'unique':
                        for col in ['Tag', 'Kassen', 'Rack']:
                            df_merge_processed.at[idx, col] = matched_row.get(col)
                            
                        # Set Brand using database value or fallback
                        if detected_brand:
                            df_merge_processed.at[idx, 'Brand'] = detected_brand
                        else:
                            brand_val = matched_row.get('Brand') if 'Brand' in matched_row else matched_row.get('brand')
                            df_merge_processed.at[idx, 'Brand'] = brand_val
                            
                        sku_val = matched_row.get('sku') if 'sku' in matched_row else matched_row.get('SKU')
                        df_merge_processed.at[idx, 'sku'] = sku_val
                        
                        inv_val = None
                        for k in ['inv', 'stock', 'stock_qty', 'qty']:
                            if k in matched_row:
                                inv_val = matched_row.get(k)
                                break
                        df_merge_processed.at[idx, 'inv'] = inv_val
                        
                        tc_val = None
                        for k in ['total cost', 'total_cost', 'totalcost']:
                            if k in matched_row:
                                tc_val = matched_row.get(k)
                                break
                        df_merge_processed.at[idx, 'total cost'] = tc_val
                        
                    # Barcode logic
                    barcode_val = matched_row.get('Barcode')
                    art_no_clean = str(matched_row.get('Art No', '')).strip()
                    if art_no_clean.endswith('.0'):
                        art_no_clean = art_no_clean[:-2]
                    item_clean = str(matched_row.get('item', '')).strip()
                    
                    has_barcode = pd.notna(barcode_val) and str(barcode_val).strip() != '' and str(barcode_val).strip().lower() != 'nan'
                    
                    if has_barcode:
                        barcode_str = str(barcode_val).strip()
                        if barcode_str.endswith('.0'):
                            barcode_str = barcode_str[:-2]
                        # Add to cache
                        if art_no_clean not in barcode_cache:
                            barcode_cache[art_no_clean] = {}
                        barcode_cache[art_no_clean][item_clean] = barcode_str
                        
                        df_merge_processed.at[idx, 'Barcode'] = barcode_str
                    else:
                        df_merge_processed.at[idx, 'Barcode'] = None
                        
            # Second pass: fill in barcode from cache for repeat rows
            for idx in range(len(df_merge_processed)):
                row_merge = df_merge_processed.iloc[idx]
                barcode_val = row_merge.get('Barcode')
                
                if pd.isna(barcode_val) or str(barcode_val).strip() == '' or str(barcode_val).strip().lower() == 'nan':
                    item_no_val = get_row_item_no(row_merge)
                    desc_val = get_row_desc(row_merge)
                    
                    cached_barcode = None
                    if item_no_val in barcode_cache and desc_val in barcode_cache[item_no_val]:
                        cached_barcode = barcode_cache[item_no_val][desc_val]
                        
                    if cached_barcode:
                        df_merge_processed.at[idx, 'Barcode'] = cached_barcode

            # Create Statistics DataFrame
            stats_data = []
            stats_data.append(("Calculation Date", pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")))
            stats_data.append(("Source Sheet Used", 'All' if self.selected_file_type != 'unique' else product_sheet_name))
            stats_data.append(("Input File Rows", len(df_original_merge)))
            stats_data.append(("Processed Rows", len(df_merge_processed)))
            
            # Find unique products in original
            item_col_orig = None
            for col in df_original_merge.columns:
                if str(col).lower() in ['item no.', 'item no', 'art no', 'art_no', 'artno']:
                    item_col_orig = col
                    break
            if item_col_orig:
                unique_orig = df_original_merge[item_col_orig].nunique()
                stats_data.append(("Input Unique Items", unique_orig))
                
            # Find unique products in processed
            item_col_proc = None
            for col in df_merge_processed.columns:
                if str(col).lower() in ['item no.', 'item no', 'art no', 'art_no', 'artno']:
                    item_col_proc = col
                    break
            if item_col_proc:
                unique_proc = df_merge_processed[item_col_proc].nunique()
                stats_data.append(("Processed Unique Items", unique_proc))
                
            # Find unique vendors in processed
            if 'Vendor' in df_merge_processed.columns:
                unique_vendors = df_merge_processed['Vendor'].nunique()
                stats_data.append(("Total Vendors Detected", unique_vendors))
                
            # Date range
            if date_col_name and date_col_name in df_merge_processed.columns:
                df_temp_dates = pd.to_datetime(df_merge_processed[date_col_name], dayfirst=True, errors='coerce').dropna()
                if not df_temp_dates.empty:
                    stats_data.append(("First Purchase Date", df_temp_dates.min().strftime("%Y-%m-%d")))
                    stats_data.append(("Last Purchase Date", df_temp_dates.max().strftime("%Y-%m-%d")))
                    
            df_stats = pd.DataFrame(stats_data, columns=["Field", "Value"])

            # Store the result and display
            if self.selected_file_type == 'unique':
                self.processed_sheets = {
                    'purchase archives': df_original_merge,
                    'unique products': df_merge_processed,
                    'Statistics': df_stats
                }
            else:
                self.processed_sheets = {
                    'purchase archives': df_original_merge,
                    'SR standard Archives': df_merge_processed,
                    'Statistics': df_stats
                }
            self._display_sheets(self.processed_sheets)
            
            # Switch to the processed sheet tab (index 1)
            self.tab_widget.setCurrentIndex(1)
            
            self.status_label.setText("Successfully applied standards. Displaying sheets.")
            self.download_btn.setVisible(True)
            self.download_btn.setEnabled(True)
            
            if self.selected_file_type == 'unique':
                self.view_tag_list_btn.setVisible(True)
                self.view_tag_list_btn.setEnabled(True)
            else:
                self.view_tag_list_btn.setVisible(False)
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.status_label.setText(f"Error applying standards: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to apply standards:\n{str(e)}")
        finally:
            self.apply_standards_btn.setEnabled(True)
            QApplication.restoreOverrideCursor()

    def download_processed(self) -> None:
        if not self.processed_sheets:
            QMessageBox.warning(self, "Error", "No processed data to download.")
            return
        
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Processed Excel", "processed_products.xlsx", "Excel Files (*.xlsx)", options=options)
        if file_name:
            try:
                with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                    for sheet_name, df in self.processed_sheets.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                QMessageBox.information(self, "Success", "File saved successfully.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save file:\n{str(e)}")



class SRVendorInvoicesDialog(QDialog):
    """Dialog for listing, rendering, parsing, and downloading vendor PDF invoices"""
    def __init__(self, parent=None, selected_vendor: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("Vendor Invoices Viewer")
        self.setMinimumSize(1200, 800)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinMaxButtonsHint)
        self.selected_vendor = selected_vendor
        self.current_pdf_path = None
        self.parsed_meta = None
        self.parsed_rows = None
        self._build_ui()
        self.load_invoices()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # Splitter for left and right panels
        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter)

        # Left panel (Invoice List)
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 5, 0)

        list_title = QLabel("Available Invoices")
        list_title.setStyleSheet("font-weight: bold; font-size: 11pt;")
        left_layout.addWidget(list_title)

        # Search bar
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Filter invoices...")
        self.search_input.textChanged.connect(self.filter_invoices)
        left_layout.addWidget(self.search_input)

        self.invoice_list = QTableWidget(0, 2)
        self.invoice_list.setHorizontalHeaderLabels(["Invoice Date", "Invoice File"])
        self.invoice_list.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.invoice_list.setSelectionBehavior(QTableWidget.SelectRows)
        self.invoice_list.setSelectionMode(QTableWidget.SingleSelection)
        self.invoice_list.itemSelectionChanged.connect(self.on_invoice_selected)
        left_layout.addWidget(self.invoice_list)

        splitter.addWidget(left_widget)

        # Right panel (Preview and Excel)
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(5, 0, 0, 0)

        # Header with actions
        header_layout = QHBoxLayout()
        self.invoice_label = QLabel("No invoice selected")
        self.invoice_label.setStyleSheet("font-weight: bold; font-size: 11pt; color: #333;")
        header_layout.addWidget(self.invoice_label, 1)

        self.parse_btn = QPushButton("To Excel")
        self.parse_btn.setStyleSheet("background-color: #28a745; color: white; padding: 6px 12px; font-weight: bold; border-radius: 4px;")
        self.parse_btn.clicked.connect(self.parse_selected_invoice)
        self.parse_btn.setEnabled(False)
        header_layout.addWidget(self.parse_btn)

        self.download_btn = QPushButton("Download Excel")
        self.download_btn.setStyleSheet("background-color: #007bff; color: white; padding: 6px 12px; font-weight: bold; border-radius: 4px;")
        self.download_btn.clicked.connect(self.download_excel)
        self.download_btn.setEnabled(False)
        header_layout.addWidget(self.download_btn)

        # Add Purchase Archives and All Unique Products buttons to the toolbar
        self.archive_btn = QPushButton("Purchase Archives")
        self.archive_btn.setStyleSheet("background-color: #17a2b8; color: white; padding: 6px 12px; font-weight: bold; border-radius: 4px;")
        self.archive_btn.clicked.connect(self.show_purchase_archives)
        self.archive_btn.setEnabled(True)
        header_layout.addWidget(self.archive_btn)

        self.all_unique_btn = QPushButton("All Unique Products")
        self.all_unique_btn.setStyleSheet("background-color: #6f42c1; color: white; padding: 6px 12px; font-weight: bold; border-radius: 4px;")
        self.all_unique_btn.clicked.connect(self.show_all_unique_products)
        self.all_unique_btn.setEnabled(True)
        header_layout.addWidget(self.all_unique_btn)

        self.upload_btn = QPushButton("Upload Invoice")
        self.upload_btn.setStyleSheet("background-color: #ffc107; color: black; padding: 6px 12px; font-weight: bold; border-radius: 4px;")
        self.upload_btn.clicked.connect(self.upload_invoice)
        self.upload_btn.setEnabled(True)
        header_layout.addWidget(self.upload_btn)

        right_layout.addLayout(header_layout)

        # Main display tabs (PDF View vs Excel View)
        self.tabs = QTabWidget()
        
        # PDF Preview Tab
        self.pdf_tab = QWidget()
        pdf_layout = QVBoxLayout(self.pdf_tab)
        pdf_layout.setContentsMargins(0, 5, 0, 0)
        
        self.pdf_scroll = QScrollArea()
        self.pdf_scroll.setWidgetResizable(True)
        self.pdf_scroll.setStyleSheet("background-color: #e0e0e0; border: 1px solid #ccc; border-radius: 4px;")
        
        # Initial empty placeholder
        self.pdf_placeholder = QLabel("Select an invoice from the list to preview")
        self.pdf_placeholder.setAlignment(Qt.AlignCenter)
        self.pdf_placeholder.setStyleSheet("color: #777; font-size: 12pt;")
        self.pdf_scroll.setWidget(self.pdf_placeholder)
        
        pdf_layout.addWidget(self.pdf_scroll)
        self.tabs.addTab(self.pdf_tab, "PDF Preview")

        # Parsed Excel Data Tab
        self.excel_tab = QWidget()
        excel_layout = QVBoxLayout(self.excel_tab)
        excel_layout.setContentsMargins(0, 5, 0, 0)

        self.excel_tabs = QTabWidget()
        
        # Sub tab 1: Invoice Items Table
        self.table_items = QTableWidget()
        self.table_items.setStyleSheet("QHeaderView::section { background-color: #f2f2f2; font-weight: bold; }")
        self.excel_tabs.addTab(self.table_items, "Invoice Data")

        # Sub tab 2: Statistics/Header Metadata
        self.table_meta = QTableWidget()
        self.table_meta.setStyleSheet("QHeaderView::section { background-color: #f2f2f2; font-weight: bold; }")
        self.excel_tabs.addTab(self.table_meta, "Statistics")

        # Sub tab 3: Purchase Archives Table
        self.table_archive = QTableWidget()
        self.table_archive.setStyleSheet("QHeaderView::section { background-color: #f2f2f2; font-weight: bold; }")
        self.excel_tabs.addTab(self.table_archive, "Purchase Archives")

        # Sub tab 3.5: Invoices Summary Table
        self.table_summary = QTableWidget()
        self.table_summary.setStyleSheet("QHeaderView::section { background-color: #f2f2f2; font-weight: bold; }")
        self.excel_tabs.addTab(self.table_summary, "Invoices Summary")

        # Sub tab 4: Unique Products Table
        self.table_unique = QTableWidget()
        self.table_unique.setStyleSheet("QHeaderView::section { background-color: #f2f2f2; font-weight: bold; }")
        self.excel_tabs.addTab(self.table_unique, "Unique Products")

        excel_layout.addWidget(self.excel_tabs)
        self.tabs.addTab(self.excel_tab, "Parsed Excel Data")
        self.tabs.setTabEnabled(1, False)  # Disabled until parsed

        right_layout.addWidget(self.tabs)

        # Status footer
        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("color: #555; padding-top: 5px;")
        layout.addWidget(self.status_label)

        # Splitter sizing
        splitter.addWidget(right_widget)
        splitter.setSizes([280, 920]) # ~1/4 to 3/4 ratio

    def extract_date_fast(self, pdf_path: Path, vendor_name: str) -> str:
        vendor_lower = vendor_name.lower()
        workspace_root = Path(r"c:\Users\mdtou\PycharmProjects\sr-vendor-handlers")
        
        # Build processed folder path
        processed_dir = workspace_root / f"processed_{vendor_lower}"
        if vendor_lower == "gft":
            processed_dir = workspace_root / "processed"
            
        processed_excel = processed_dir / f"{pdf_path.stem}.xlsx"
        
        # 1. Try to read from processed Excel Statistics sheet
        if processed_excel.exists():
            try:
                df_stat = pd.read_excel(processed_excel, sheet_name="Statistics")
                if not df_stat.empty:
                    for c in ["invoice date", "datum", "date"]:
                        for col in df_stat.columns:
                            if str(col).lower() == c:
                                return str(df_stat.iloc[0][col])
            except Exception:
                pass
                
        # 2. Fallback: Parse the first page text of the PDF directly using pypdfium2 (very fast)
        try:
            doc = pdfium.PdfDocument(str(pdf_path))
            if len(doc) > 0:
                page = doc[0]
                textpage = page.get_textpage()
                text = textpage.get_text_bounded()
                # Find date pattern (DD.MM.YYYY, DD-MM-YYYY, DD/MM/YYYY, etc.)
                matches = re.findall(r"\b\d{1,2}[-./]\d{1,2}[-./]\d{2,4}\b", text)
                if matches:
                    return matches[0]
        except Exception:
            pass
            
        return "N/A"

    def upload_invoice(self) -> None:
        if not self.selected_vendor:
            QMessageBox.warning(self, "No Vendor", "Please select a vendor first.")
            return

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Upload Invoice PDF",
            "",
            "PDF Files (*.pdf)",
            options=options
        )
        if not file_path:
            return

        try:
            src_path = Path(file_path)
            # Destination directory: data/vendors/<vendor_name>/invoices/
            dest_dir = VENDOR_ROOT / self.selected_vendor.lower() / "invoices"
            dest_dir.mkdir(parents=True, exist_ok=True)
            
            dest_path = dest_dir / src_path.name
            
            # If a file with the same name already exists in destination, verify overwrite
            if dest_path.exists():
                reply = QMessageBox.question(
                    self,
                    "Overwrite File?",
                    f"An invoice with the name '{src_path.name}' already exists.\nDo you want to overwrite it?",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply != QMessageBox.Yes:
                    return

            # Read source bytes and write to destination
            dest_path.write_bytes(src_path.read_bytes())
            
            QMessageBox.information(
                self,
                "Success",
                f"Successfully uploaded and saved invoice to:\n{dest_path}"
            )
            
            # Reload invoices list to show new file
            self.load_invoices()
        except Exception as e:
            QMessageBox.critical(
                self,
                "Upload Error",
                f"Failed to upload invoice file:\n{str(e)}"
            )

    def load_invoices(self) -> None:
        self.invoice_list.setRowCount(0)
        if not VENDOR_ROOT.exists():
            self.status_label.setText("Error: Vendor data directory not found.")
            return

        pdf_files = list(VENDOR_ROOT.glob("*/invoices/*.pdf"))
        
        # Filter files if selected_vendor is specified
        if self.selected_vendor:
            pdf_files = [p for p in pdf_files if p.parent.parent.name.lower() == self.selected_vendor.lower()]

        if not pdf_files:
            msg = f"No PDF invoices found for vendor '{self.selected_vendor}'." if self.selected_vendor else "No PDF invoices found in vendor directories."
            self.status_label.setText(msg)
            self.invoice_list.setRowCount(1)
            item = QTableWidgetItem("No invoices found")
            item.setFlags(Qt.NoItemFlags)
            self.invoice_list.setItem(0, 0, item)
            self.invoice_list.setItem(0, 1, QTableWidgetItem(""))
            return

        # Collect invoice info
        invoices_info = []
        for pdf_path in pdf_files:
            vendor_name = pdf_path.parent.parent.name.capitalize()
            file_display = f"[{vendor_name}] {pdf_path.name}"
            invoice_date = self.extract_date_fast(pdf_path, vendor_name)
            
            # Parse date to datetime object for sorting descending
            if invoice_date != "N/A":
                parsed_date = pd.to_datetime(invoice_date, dayfirst=True, errors="coerce")
                if pd.isna(parsed_date):
                    parsed_date = pd.Timestamp.min
            else:
                parsed_date = pd.Timestamp.min
                
            invoices_info.append((parsed_date, pdf_path, file_display, invoice_date))

        # Sort chronologically descending (latest date first)
        invoices_info.sort(key=lambda x: x[0], reverse=True)

        self.invoice_list.setRowCount(len(invoices_info))
        for idx, (_, pdf_path, file_display, invoice_date) in enumerate(invoices_info):
            item_date = QTableWidgetItem(invoice_date)
            item_date.setTextAlignment(Qt.AlignCenter)
            
            item_file = QTableWidgetItem(file_display)
            item_file.setData(Qt.UserRole, str(pdf_path))
            
            # Swapped columns: Date is column 0, File is column 1
            self.invoice_list.setItem(idx, 0, item_date)
            self.invoice_list.setItem(idx, 1, item_file)

        self.invoice_list.resizeColumnsToContents()
        self.status_label.setText(f"Found {len(pdf_files)} PDF invoices.")

    def filter_invoices(self) -> None:
        search_text = self.search_input.text().lower()
        for i in range(self.invoice_list.rowCount()):
            item = self.invoice_list.item(i, 1) # Column 1 contains File Name and UserRole
            if item and item.data(Qt.UserRole):
                self.invoice_list.setRowHidden(i, search_text not in item.text().lower())

    def on_invoice_selected(self) -> None:
        selected_ranges = self.invoice_list.selectedRanges()
        if not selected_ranges:
            return
            
        row = selected_ranges[0].topRow()
        item = self.invoice_list.item(row, 1) # Column 1 contains File Name and UserRole
        if not item:
            return
            
        pdf_path_str = item.data(Qt.UserRole)
        if not pdf_path_str:
            return

        self.current_pdf_path = Path(pdf_path_str)
        vendor_name = self.current_pdf_path.parent.parent.name
        
        self.invoice_label.setText(f"{vendor_name.capitalize()} Invoice: {self.current_pdf_path.name}")
        self.status_label.setText(f"Loading preview for {self.current_pdf_path.name}...")
        
        # Reset parsed data
        self.parsed_meta = None
        self.parsed_rows = None
        self.archive_rows = []
        self.unique_rows = []
        self.summary_rows = []
        self.tabs.setTabEnabled(1, False)
        self.download_btn.setEnabled(False)
        self.tabs.setCurrentIndex(0)

        # Check if parser is available
        parser = get_parser(vendor_name)
        if parser is not None:
            self.parse_btn.setEnabled(True)
            self.parse_btn.setToolTip("Click to convert PDF invoice to Excel data")
            self.parse_btn.setStyleSheet("background-color: #28a745; color: white; padding: 6px 12px; font-weight: bold; border-radius: 4px;")
        else:
            self.parse_btn.setEnabled(False)
            self.parse_btn.setToolTip(f"No parser available for vendor '{vendor_name}' yet.")
            self.parse_btn.setStyleSheet("background-color: #6c757d; color: white; padding: 6px 12px; font-weight: bold; border-radius: 4px;")

        # Render PDF pages
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            doc = pdfium.PdfDocument(str(self.current_pdf_path))
            
            container = QWidget()
            container.setStyleSheet("background-color: #e0e0e0;")
            scroll_layout = QVBoxLayout(container)
            scroll_layout.setSpacing(15)
            scroll_layout.setContentsMargins(10, 10, 10, 10)
            
            for i in range(len(doc)):
                page = doc[i]
                # Render to PIL Image
                pil_img = page.render(scale=1.5).to_pil()
                
                # Convert PIL Image to QPixmap via BytesIO
                buf = io.BytesIO()
                pil_img.save(buf, format='PNG')
                qimg = QImage()
                qimg.loadFromData(buf.getvalue(), 'PNG')
                pixmap = QPixmap.fromImage(qimg)
                
                label = QLabel()
                label.setPixmap(pixmap)
                label.setAlignment(Qt.AlignCenter)
                label.setStyleSheet("border: 1px solid #bbb; background-color: white;")
                scroll_layout.addWidget(label)
                
            self.pdf_scroll.setWidget(container)
            self.status_label.setText(f"Loaded {self.current_pdf_path.name} ({len(doc)} pages).")
        except Exception as e:
            self.status_label.setText(f"Failed to render preview: {str(e)}")
            placeholder = QLabel(f"Failed to render PDF preview:\n{str(e)}")
            placeholder.setAlignment(Qt.AlignCenter)
            placeholder.setStyleSheet("color: red; font-size: 11pt;")
            self.pdf_scroll.setWidget(placeholder)
        finally:
            QApplication.restoreOverrideCursor()

    def get_merged_history_path(self) -> Optional[Path]:
        if not self.selected_vendor:
            return None
            
        vendor_lower = self.selected_vendor.lower()
        workspace_root = Path(r"c:\Users\mdtou\PycharmProjects\sr-vendor-handlers")
        
        # Candidate 1: merged_<vendor>.xlsx
        path = workspace_root / f"merged_{vendor_lower}.xlsx"
        if not path.exists() and vendor_lower == "gft":
            # Candidate 2: merged.xlsx
            path = workspace_root / "merged.xlsx"
            
        # Try general scanning
        if not path.exists():
            for f in workspace_root.glob("merged_*.xlsx"):
                if vendor_lower in f.name.lower():
                    path = f
                    break
                    
        return path if path.exists() else None

    def load_merged_history(self) -> Optional[pd.DataFrame]:
        path = self.get_merged_history_path()
        if not path:
            return None
            
        try:
            df = pd.read_excel(path, sheet_name="Invoice Data")
            return df
        except Exception as e:
            print(f"Error loading merged history from {path}: {e}")
            return None

    def get_product_id_column(self, df: pd.DataFrame) -> Optional[str]:
        # Common item code column names
        candidates = ["item no.", "artnr", "article", "code", "article number", "itemcode"]
        for col in df.columns:
            if str(col).lower() in candidates:
                return col
        return df.columns[0] if len(df.columns) > 0 else None

    def get_amount_column(self, df: pd.DataFrame) -> Optional[str]:
        # Common amount column names
        candidates = ["amount", "betrag eur", "net price", "price"]
        for col in df.columns:
            if str(col).lower() in candidates:
                return col
        return None

    def show_purchase_archives(self) -> None:
        if not self.selected_vendor:
            QMessageBox.warning(self, "No Vendor", "Please select a vendor first.")
            return
            
        df = self.load_merged_history()
        if df is None or df.empty:
            QMessageBox.information(self, "No Archive", f"No purchase archive (merged Excel file) found for vendor '{self.selected_vendor}'.")
            return
            
        # Convert df to rows list of dicts to use our helper
        rows = df.to_dict(orient="records")
        self.archive_rows = rows
        
        # Populate table_archive
        self.populate_table(self.table_archive, rows)
        
        # Calculate stats for the archive
        date_col = None
        for c in ["invoice date", "datum", "date"]:
            for actual_col in df.columns:
                if str(actual_col).lower() == c:
                    date_col = actual_col
                    break
            if date_col:
                break
        
        stats_meta = {
            "Total Archived Products": len(rows),
        }
        
        inv_col = None
        for c in ["invoice number", "rechn.nr.", "doc. nr.", "invoice_no", "invoice no"]:
            for actual_col in df.columns:
                if str(actual_col).lower() == c:
                    inv_col = actual_col
                    break
            if inv_col:
                break
                
        if inv_col:
            stats_meta["Total Distinct Invoices"] = df[inv_col].nunique()
            
        if date_col and not df.empty:
            df_sorted = df.copy()
            df_sorted["_parsed_date"] = pd.to_datetime(df_sorted[date_col], dayfirst=True, errors="coerce")
            df_sorted = df_sorted.sort_values(by="_parsed_date", ascending=True)
            stats_meta["First Purchase Date"] = df_sorted[date_col].iloc[0]
            stats_meta["Last Purchase Date"] = df_sorted[date_col].iloc[-1]
            
        self.populate_meta_table(self.table_meta, stats_meta)
        self.parsed_meta = stats_meta
        
        # Load Statistics sheet from history to get correct total amounts
        path = self.get_merged_history_path()
        df_stat = None
        if path:
            try:
                df_stat = pd.read_excel(path, sheet_name="Statistics")
            except Exception as e:
                print(f"Error loading Statistics sheet: {e}")

        # Find columns in Statistics sheet
        stat_inv_col = "Invoice Number"
        stat_date_col = "Invoice Date"
        stat_amt_col = "Total Amount"
        
        if df_stat is not None and not df_stat.empty:
            for col in df_stat.columns:
                if str(col).lower() in ["rechn.nr.", "invoice number", "invoice_no"]:
                    stat_inv_col = col
                if str(col).lower() in ["datum", "invoice date", "date"]:
                    stat_date_col = col
                if str(col).lower() in ["endbetrag eur", "total amount", "amount"]:
                    stat_amt_col = col

        summary_rows = []
        if inv_col:
            df_chron = df.copy()
            if date_col:
                df_chron["_parsed_date"] = pd.to_datetime(df_chron[date_col], dayfirst=True, errors="coerce")
                df_chron = df_chron.sort_values(by="_parsed_date", ascending=True)
            
            invoice_pids = {}
            id_col = self.get_product_id_column(df_chron)
            for _, r in df_chron.iterrows():
                inv_no = str(r[inv_col]).strip()
                if inv_no not in invoice_pids:
                    invoice_pids[inv_no] = []
                if id_col:
                    invoice_pids[inv_no].append(str(r[id_col]).strip())

            # If Statistics sheet was loaded, use its values
            if df_stat is not None and not df_stat.empty:
                df_stat_sorted = df_stat.copy()
                df_stat_sorted["_parsed_date"] = pd.to_datetime(df_stat_sorted[stat_date_col], dayfirst=True, errors="coerce")
                df_stat_sorted = df_stat_sorted.sort_values(by="_parsed_date", ascending=True)
                
                seen_pids = set()
                for _, row_stat in df_stat_sorted.iterrows():
                    inv_no = str(row_stat[stat_inv_col]).strip()
                    inv_date = str(row_stat[stat_date_col]).strip()
                    total_amt = float(row_stat[stat_amt_col]) if not pd.isna(row_stat[stat_amt_col]) else 0.0
                    
                    pids_in_inv = invoice_pids.get(inv_no, [])
                    prod_count = len(pids_in_inv)
                    if "product_count" in row_stat:
                        try:
                            prod_count = int(row_stat["product_count"])
                        except (ValueError, TypeError):
                            pass
                            
                    unique_count = 0
                    for pid in pids_in_inv:
                        if pid not in seen_pids:
                            seen_pids.add(pid)
                            unique_count += 1
                            
                    summary_rows.append({
                        "Invoice Number": inv_no,
                        "Invoice Date": inv_date,
                        "Total Amount": round(total_amt, 2),
                        "Ordered Product Count": prod_count,
                        "Unique Products Count": unique_count
                    })
            else:
                # Fallback to row summing if Statistics sheet is missing/empty
                invoice_groups = {}
                for _, r in df_chron.iterrows():
                    inv_no = str(r[inv_col]).strip()
                    inv_date = str(r[date_col]).strip() if date_col else "N/A"
                    key = (inv_no, inv_date)
                    if key not in invoice_groups:
                        invoice_groups[key] = []
                    invoice_groups[key].append(r)
                    
                seen_pids = set()
                amount_col = self.get_amount_column(df_chron)
                
                for (inv_no, inv_date), group_rows in invoice_groups.items():
                    total_amt = 0.0
                    unique_count = 0
                    for r in group_rows:
                        if amount_col:
                            try:
                                val = float(r[amount_col])
                                if not pd.isna(val):
                                    total_amt += val
                            except (ValueError, TypeError):
                                pass
                        if id_col:
                            pid = str(r[id_col]).strip()
                            if pid not in seen_pids:
                                seen_pids.add(pid)
                                unique_count += 1
                                
                    summary_rows.append({
                        "Invoice Number": inv_no,
                        "Invoice Date": inv_date,
                        "Total Amount": round(total_amt, 2),
                        "Ordered Product Count": len(group_rows),
                        "Unique Products Count": unique_count
                    })

            if date_col:
                # Sort summary descending (latest invoice date first)
                summary_rows.sort(
                    key=lambda x: pd.to_datetime(x["Invoice Date"], dayfirst=True, errors="coerce") or pd.Timestamp.min,
                    reverse=True
                )
                
        self.summary_rows = summary_rows
        self.populate_table(self.table_summary, self.summary_rows)
        
        # Setup tab visibility (Show Statistics, Purchase Archives, Invoices Summary only)
        self.excel_tabs.setTabVisible(0, False)
        self.excel_tabs.setTabVisible(1, True)
        self.excel_tabs.setTabVisible(2, True)
        self.excel_tabs.setTabVisible(3, True)
        self.excel_tabs.setTabVisible(4, False)
        
        # Enable Excel view and switch to the archive tab
        self.tabs.setTabEnabled(1, True)
        self.tabs.setCurrentIndex(1)
        self.excel_tabs.setCurrentIndex(2) # "Purchase Archives" tab is index 2
        self.download_btn.setEnabled(True)
        self.status_label.setText(f"Loaded {len(rows)} products and summarized {len(summary_rows)} invoices.")

    def show_all_unique_products(self) -> None:
        if not self.selected_vendor:
            QMessageBox.warning(self, "No Vendor", "Please select a vendor first.")
            return
            
        df = self.load_merged_history()
        if df is None or df.empty:
            QMessageBox.information(self, "No Archive", f"No purchase archive (merged Excel file) found for vendor '{self.selected_vendor}'.")
            return

        # Try parsing dates chronologically
        date_col = None
        for c in ["invoice date", "datum", "date"]:
            for actual_col in df.columns:
                if str(actual_col).lower() == c:
                    date_col = actual_col
                    break
            if date_col:
                break
                
        if date_col:
            df = df.copy()
            df["_parsed_date"] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
            df = df.sort_values(by="_parsed_date", ascending=True)

        id_col = self.get_product_id_column(df)
        if not id_col:
            QMessageBox.warning(self, "Error", "Could not find product code column in merged data.")
            return

        seen = set()
        unique_rows = []
        for _, row in df.iterrows():
            pid = str(row[id_col]).strip()
            if pid not in seen:
                seen.add(pid)
                row_dict = row.to_dict()
                if "_parsed_date" in row_dict:
                    del row_dict["_parsed_date"]
                unique_rows.append(row_dict)

        # Let's populate the unique table
        self.populate_table(self.table_unique, unique_rows)
        self.unique_rows = unique_rows
        
        # Let's generate statistics
        stats_meta = {
            "Total Unique Products": len(unique_rows),
        }
        
        # Distinct invoices
        inv_col = None
        for c in ["invoice number", "rechn.nr.", "doc. nr.", "invoice_no", "invoice no"]:
            for actual_col in df.columns:
                if str(actual_col).lower() == c:
                    inv_col = actual_col
                    break
            if inv_col:
                break
                
        if inv_col:
            stats_meta["Total Distinct Invoices"] = df[inv_col].nunique()
            
        if date_col and not df.empty:
            stats_meta["First Purchase Date"] = df[date_col].iloc[0]
            stats_meta["Last Purchase Date"] = df[date_col].iloc[-1]
            
        self.populate_meta_table(self.table_meta, stats_meta)
        self.parsed_meta = stats_meta
        self.summary_rows = []
        
        # Setup tab visibility (Show Statistics and Unique Products only)
        self.excel_tabs.setTabVisible(0, False)
        self.excel_tabs.setTabVisible(1, True)
        self.excel_tabs.setTabVisible(2, False)
        self.excel_tabs.setTabVisible(3, False)
        self.excel_tabs.setTabVisible(4, True)
        
        # Switch tabs
        self.tabs.setTabEnabled(1, True)
        self.tabs.setCurrentIndex(1)
        self.excel_tabs.setCurrentIndex(4) # "Unique Products" tab is index 4
        self.download_btn.setEnabled(True)
        self.status_label.setText(f"Computed {len(unique_rows)} total unique products in history.")

    def get_meta_inv_details(self, meta: Dict[str, Any]) -> Tuple[str, str]:
        date_val = ""
        inv_val = ""
        
        # Common date keys
        for k in ["invoice date", "datum", "date"]:
            for actual_key in meta.keys():
                if str(actual_key).lower() == k:
                    date_val = str(meta[actual_key])
                    break
            if date_val:
                break
                
        # Common invoice number keys
        for k in ["invoice number", "rechn.nr.", "doc. nr.", "invoice_no", "invoice no"]:
            for actual_key in meta.keys():
                if str(actual_key).lower() == k:
                    inv_val = str(meta[actual_key])
                    break
            if inv_val:
                break
                
        return date_val, inv_val

    def get_prior_history(self, df_archive: pd.DataFrame, current_date_str: str, current_inv_no: str) -> pd.DataFrame:
        if df_archive.empty:
            return df_archive
            
        df = df_archive.copy()
        
        # Find date column
        date_col = None
        for c in ["invoice date", "datum", "date"]:
            for actual_col in df.columns:
                if str(actual_col).lower() == c:
                    date_col = actual_col
                    break
            if date_col:
                break
                
        # Find invoice number column
        inv_col = None
        for c in ["invoice number", "rechn.nr.", "doc. nr.", "invoice_no", "invoice no"]:
            for actual_col in df.columns:
                if str(actual_col).lower() == c:
                    inv_col = actual_col
                    break
            if inv_col:
                break
                
        # Parse dates in archive
        if date_col:
            df["_parsed_date"] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
            # Sort chronologically
            df = df.sort_values(by="_parsed_date", ascending=True)
            
        # Parse current date
        current_date = pd.to_datetime(current_date_str, dayfirst=True, errors="coerce") if current_date_str else None
        
        # If invoice number is found in the archive, filter to keep all rows before its first occurrence
        if inv_col and current_inv_no:
            matching_indices = df[df[inv_col].astype(str) == str(current_inv_no)].index
            if len(matching_indices) > 0:
                # Get the position of the first matching row in the sorted DataFrame
                sorted_indices = list(df.index)
                first_match_pos = sorted_indices.index(matching_indices[0])
                # Prior rows are all those before first_match_pos
                prior_df = df.iloc[:first_match_pos]
                return prior_df
                
        # Fallback: if invoice number is not in the archive, filter by date strictly less than current_date
        if date_col and current_date is not None:
            prior_df = df[df["_parsed_date"] < current_date]
            return prior_df
            
        # Fallback 2: if date parsing fails or no dates, return empty DataFrame (assume everything is unique)
        return pd.DataFrame(columns=df_archive.columns)

    def parse_selected_invoice(self) -> None:
        if not self.current_pdf_path:
            return

        vendor_name = self.current_pdf_path.parent.parent.name
        parser = get_parser(vendor_name)
        if not parser:
            QMessageBox.warning(self, "No Parser", f"No parser registered for vendor '{vendor_name}'.")
            return

        self.status_label.setText(f"Parsing invoice using {vendor_name.capitalize()} parser...")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            # Parse the PDF using modular parser
            meta, rows = parser.parse(self.current_pdf_path)
            
            # Save the parsed invoice to the processed directory
            workspace_root = Path(r"c:\Users\mdtou\PycharmProjects\sr-vendor-handlers")
            vendor_lower = vendor_name.lower()
            processed_dir = workspace_root / f"processed_{vendor_lower}"
            if vendor_lower == "gft":
                processed_dir = workspace_root / "processed"
            processed_dir.mkdir(parents=True, exist_ok=True)
            
            excel_out_path = processed_dir / f"{self.current_pdf_path.stem}.xlsx"
            parser.write_excel(meta, rows, excel_out_path)
            
            # Regenerate the merged Excel file
            self.regenerate_merged_file(vendor_lower, processed_dir)
            
            # Load historical merged data (now updated with the parsed invoice)
            df_archive = self.load_merged_history()
            
            unique_current_rows = []
            
            if df_archive is not None and not df_archive.empty:
                # Find date & invoice number of current selected invoice from meta
                curr_date_str, curr_inv_no = self.get_meta_inv_details(meta)
                
                # Retrieve history strictly prior to the current selected invoice
                df_prior = self.get_prior_history(df_archive, curr_date_str, curr_inv_no)
                
                # Find product code column in current rows and prior archive df
                id_col_current = None
                if rows:
                    id_col_current = self.get_product_id_column(pd.DataFrame(rows[:1]))
                    
                id_col_archive = self.get_product_id_column(df_prior)
                
                if id_col_current and id_col_archive and not df_prior.empty:
                    prior_ids = set(str(pid).strip() for pid in df_prior[id_col_archive].dropna())
                    for r in rows:
                        pid = str(r.get(id_col_current)).strip()
                        if pid not in prior_ids:
                            unique_current_rows.append(r)
                else:
                    unique_current_rows = rows
            else:
                unique_current_rows = rows
                
            meta["Total Ordered Product Count"] = len(rows)
            meta["Unique Products Count"] = len(unique_current_rows)
            
            self.parsed_meta = meta
            self.parsed_rows = rows
            self.archive_rows = []
            self.unique_rows = unique_current_rows
            self.summary_rows = []
            
            # Populate tables in Excel tab
            self.populate_table(self.table_items, self.parsed_rows)
            self.populate_meta_table(self.table_meta, self.parsed_meta)
            self.populate_table(self.table_unique, self.unique_rows)
            
            # Setup tab visibility (Hide Purchase Archives and Invoices Summary for individual invoices)
            self.excel_tabs.setTabVisible(0, True)
            self.excel_tabs.setTabVisible(1, True)
            self.excel_tabs.setTabVisible(2, False)
            self.excel_tabs.setTabVisible(3, False)
            self.excel_tabs.setTabVisible(4, True)
            
            # Enable Excel views
            self.tabs.setTabEnabled(1, True)
            self.tabs.setCurrentIndex(1)  # Automatically switch to excel preview
            self.excel_tabs.setCurrentIndex(0) # Default to Invoice Data
            self.download_btn.setEnabled(True)
            self.status_label.setText("Successfully parsed invoice to Excel data.")
            
        except Exception as e:
            self.status_label.setText(f"Parsing failed: {str(e)}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Parsing Error", f"Failed to parse invoice:\n{str(e)}")
        finally:
            QApplication.restoreOverrideCursor()

    def regenerate_merged_file(self, vendor_lower: str, processed_dir: Path) -> None:
        try:
            xlsx_files = list(processed_dir.glob("*.xlsx"))
            if not xlsx_files:
                return
                
            all_master_rows = []
            all_summary_rows = []
            
            for excel_path in xlsx_files:
                try:
                    df_data = pd.read_excel(excel_path, sheet_name="Invoice Data")
                    df_stat = pd.read_excel(excel_path, sheet_name="Statistics")
                    
                    if not df_data.empty:
                        all_master_rows.append(df_data)
                    if not df_stat.empty:
                        all_summary_rows.append(df_stat)
                except Exception as e:
                    print(f"Error reading processed file {excel_path} during merge: {e}")
                    
            if all_master_rows:
                df = pd.concat(all_master_rows, ignore_index=True)
                df_summary = pd.concat(all_summary_rows, ignore_index=True)
                
                # Determine date column for sorting
                date_col = None
                for c in ["invoice date", "datum", "date"]:
                    for actual_col in df.columns:
                        if str(actual_col).lower() == c:
                            date_col = actual_col
                            break
                    if date_col:
                        break
                        
                # Determine summary date column for sorting
                sum_date_col = None
                for c in ["invoice date", "datum", "date"]:
                    for actual_col in df_summary.columns:
                        if str(actual_col).lower() == c:
                            sum_date_col = actual_col
                            break
                    if sum_date_col:
                        break
                
                if date_col:
                    df["Datum_parsed"] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
                    df = df.sort_values(by="Datum_parsed", ascending=False)
                    df = df.drop(columns=["Datum_parsed"])
                    
                if sum_date_col:
                    df_summary["Datum_parsed"] = pd.to_datetime(df_summary[sum_date_col], dayfirst=True, errors="coerce")
                    df_summary = df_summary.sort_values(by="Datum_parsed", ascending=False)
                    df_summary = df_summary.drop(columns=["Datum_parsed"])
                    
                workspace_root = Path(r"c:\Users\mdtou\PycharmProjects\sr-vendor-handlers")
                merged_path = workspace_root / f"merged_{vendor_lower}.xlsx"
                if vendor_lower == "gft":
                    merged_path = workspace_root / "merged.xlsx"
                    
                with pd.ExcelWriter(merged_path, engine="openpyxl") as writer:
                    df.to_excel(writer, sheet_name="Invoice Data", index=False)
                    df_summary.to_excel(writer, sheet_name="Statistics", index=False)
                    
                print(f"Successfully regenerated merged history for {vendor_lower} at {merged_path}")
        except Exception as e:
            print(f"Error regenerating merged file: {e}")

    def populate_table(self, table: QTableWidget, data: List[Dict[str, Any]]) -> None:
        table.clear()
        if not data:
            table.setRowCount(0)
            table.setColumnCount(0)
            return
            
        headers = list(data[0].keys())
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setRowCount(len(data))
        
        for row_idx, row_data in enumerate(data):
            for col_idx, key in enumerate(headers):
                val = row_data.get(key)
                if isinstance(val, float):
                    item = QTableWidgetItem(f"{val:.2f}")
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    item = QTableWidgetItem(str(val) if val is not None else "")
                    if isinstance(val, int):
                        item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    else:
                        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                table.setItem(row_idx, col_idx, item)
        table.resizeColumnsToContents()

    def populate_meta_table(self, table: QTableWidget, meta: Dict[str, Any]) -> None:
        table.clear()
        table.setColumnCount(2)
        table.setHorizontalHeaderLabels(["Field", "Value"])
        table.setRowCount(len(meta))
        
        for idx, (key, val) in enumerate(meta.items()):
            table.setItem(idx, 0, QTableWidgetItem(str(key)))
            val_str = f"{val:.2f}" if isinstance(val, float) else str(val)
            item = QTableWidgetItem(val_str if val is not None else "")
            table.setItem(idx, 1, item)
        table.resizeColumnsToContents()

    def download_excel(self) -> None:
        # Check if we have visible data to download
        has_data = False
        if self.excel_tabs.isTabVisible(0) and getattr(self, "parsed_rows", None):
            has_data = True
        if self.excel_tabs.isTabVisible(1) and getattr(self, "parsed_meta", None):
            has_data = True
        if self.excel_tabs.isTabVisible(2) and getattr(self, "archive_rows", None):
            has_data = True
        if self.excel_tabs.isTabVisible(3) and getattr(self, "summary_rows", None):
            has_data = True
        if self.excel_tabs.isTabVisible(4) and getattr(self, "unique_rows", None):
            has_data = True

        if not has_data:
            QMessageBox.warning(self, "No Data", "There is no visible data to export.")
            return

        if self.current_pdf_path:
            default_name = f"{self.current_pdf_path.stem}_parsed.xlsx"
        elif self.selected_vendor:
            if self.excel_tabs.isTabVisible(2):
                default_name = f"{self.selected_vendor.lower()}_purchase_archives.xlsx"
            else:
                default_name = f"{self.selected_vendor.lower()}_unique_products.xlsx"
        else:
            default_name = "export.xlsx"

        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Parsed Excel", default_name, "Excel Files (*.xlsx)", options=options)
        if file_name:
            self.status_label.setText(f"Saving Excel to {file_name}...")
            QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                sheets_written = []
                with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                    # Sheet 1: Invoice Data (Index 0)
                    if self.excel_tabs.isTabVisible(0) and getattr(self, "parsed_rows", None):
                        pd.DataFrame(self.parsed_rows).to_excel(writer, sheet_name="Invoice Data", index=False)
                        sheets_written.append("Invoice Data")
                        
                    # Sheet 2: Statistics (Index 1)
                    if self.excel_tabs.isTabVisible(1) and getattr(self, "parsed_meta", None):
                        meta_df = pd.DataFrame(list(self.parsed_meta.items()), columns=["Field", "Value"])
                        meta_df.to_excel(writer, sheet_name="Statistics", index=False)
                        sheets_written.append("Statistics")
                        
                    # Sheet 3: Purchase Archives (Index 2)
                    if self.excel_tabs.isTabVisible(2):
                        archive_rows = getattr(self, "archive_rows", [])
                        if archive_rows:
                            pd.DataFrame(archive_rows).to_excel(writer, sheet_name="Purchase Archives", index=False)
                            sheets_written.append("Purchase Archives")
                            
                    # Sheet 4: Invoices Summary (Index 3)
                    if self.excel_tabs.isTabVisible(3):
                        summary_rows = getattr(self, "summary_rows", [])
                        if summary_rows:
                            pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Invoices Summary", index=False)
                            sheets_written.append("Invoices Summary")
                            
                    # Sheet 5: Unique Products (Index 4)
                    if self.excel_tabs.isTabVisible(4):
                        unique_rows = getattr(self, "unique_rows", [])
                        if unique_rows:
                            pd.DataFrame(unique_rows).to_excel(writer, sheet_name="Unique Products", index=False)
                            sheets_written.append("Unique Products")
                            
                self.status_label.setText(f"Excel saved to {file_name}")
                sheets_str = ", ".join(sheets_written)
                QMessageBox.information(self, "Success", f"Invoice successfully parsed and Excel saved with sheets: {sheets_str} to:\n{file_name}")
            except Exception as e:
                self.status_label.setText(f"Failed to save Excel: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to save Excel file:\n{str(e)}")
            finally:
                QApplication.restoreOverrideCursor()


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
        init_brand_db()
        self.db = DatabaseManager(DB_FILE)
        self.loader = ExcelLoader(VENDOR_ROOT)
        self.current_vendor: Optional[str] = None
        self.vendor_names: List[str] = []
        self.product_cache: List[Dict] = []
        self._build_ui()
        self._load_sidebar()
        self._refresh_all_stats()

        # Start web server for mobile access
        try:
            from web_server import start_web_server
            url = start_web_server(self.db, lambda: self.vendor_names)
            if "cryptography" in url:
                real_url = url.split(" (")[0]
                self.server_info_label.setText(f"<b>Mobile Server:</b><br><a href='{real_url}'>{real_url}</a><br><small style='color:red;'>Run <b>pip install cryptography</b> for camera!</small>")
            else:
                self.server_info_label.setText(f"<b>Mobile Server:</b><br><a href='{url}'>{url}</a><br><small>Accept browser security warnings for local IP.</small>")
            self.server_info_label.setOpenExternalLinks(True)
        except ImportError:
            self.server_info_label.setText("Install 'flask' to enable mobile access.")

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
        self.invoices_button = QPushButton("Vendor Invoices")
        self.invoices_button.clicked.connect(self.open_invoices_dialog)
        self.reset_vendor_button = QPushButton("Reset selected vendor products")
        self.reset_vendor_button.clicked.connect(self.reset_selected_vendor_products)
        self.inventory_button = QPushButton("Product inventory")
        self.inventory_button.clicked.connect(self.open_inventory_dialog)
        self.order_button = QPushButton("Create order")
        self.order_button.clicked.connect(self.open_order_dialog)
        self.vendor_details_button = QPushButton("Vendor details")
        self.vendor_details_button.clicked.connect(self.open_vendor_details)
        self.refresh_view_button = QPushButton("Refresh product list")
        self.refresh_view_button.clicked.connect(self.refresh_product_list)
        self.archive_button = QPushButton("SR Products Archive")
        self.archive_button.clicked.connect(self.open_archive_dialog)
        self.shopfloor_brands_button = QPushButton("Shopfloor Retail Brands")
        self.shopfloor_brands_button.clicked.connect(self.open_shopfloor_brands_dialog)
        left_panel.addWidget(self.sync_vendor_button)
        left_panel.addWidget(self.sync_all_button)
        left_panel.addWidget(self.invoices_button)
        left_panel.addWidget(self.reset_vendor_button)
        left_panel.addWidget(self.inventory_button)
        left_panel.addWidget(self.order_button)
        left_panel.addWidget(self.vendor_details_button)
        left_panel.addWidget(self.refresh_view_button)
        left_panel.addWidget(self.archive_button)
        left_panel.addWidget(self.shopfloor_brands_button)
        
        left_panel.addStretch()
        self.server_info_label = QLabel("Initializing server...")
        left_panel.addWidget(self.server_info_label)

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

    def refresh_product_list(self) -> None:
        if self.current_vendor:
            self._load_products_for_vendor(self.current_vendor)
            self._refresh_all_stats()

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

    def open_archive_dialog(self) -> None:
        dialog = SRProductsArchiveDialog(parent=self)
        dialog.exec_()

    def open_shopfloor_brands_dialog(self) -> None:
        dialog = ShopfloorBrandsDialog(parent=self)
        dialog.exec_()

    def open_invoices_dialog(self) -> None:
        if not self.current_vendor:
            QMessageBox.warning(self, "No vendor selected", "Please select a vendor before opening vendor invoices.")
            return
        dialog = SRVendorInvoicesDialog(parent=self, selected_vendor=self.current_vendor)
        dialog.exec_()

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
        
        # Try to fill in missing/zero prices from merged invoice history
        self.enrich_prices_from_history(vendor, products)
        
        self.db.upsert_vendor_products(vendor, products)
        if self.current_vendor == vendor:
            self._load_products_for_vendor(vendor)
        self._refresh_all_stats()
        QMessageBox.information(self, "Sync complete", f"{len(products)} products synced for {vendor}.")

    def enrich_prices_from_history(self, vendor: str, products: List[Dict[str, Any]]) -> None:
        vendor_lower = vendor.lower()
        workspace_root = Path(r"c:\Users\mdtou\PycharmProjects\sr-vendor-handlers")
        merged_path = workspace_root / f"merged_{vendor_lower}.xlsx"
        if not merged_path.exists() and vendor_lower == "gft":
            merged_path = workspace_root / "merged.xlsx"
            
        if not merged_path.exists():
            for f in workspace_root.glob("merged_*.xlsx"):
                if vendor_lower in f.name.lower():
                    merged_path = f
                    break
                    
        if not merged_path.exists():
            return
            
        try:
            df = pd.read_excel(merged_path, sheet_name="Invoice Data")
            if df.empty:
                return
                
            # Find product ID and price columns
            id_col = None
            for c in ["item no.", "artnr", "article", "code", "article number", "itemcode"]:
                for col in df.columns:
                    if str(col).lower() == c:
                        id_col = col
                        break
                if id_col:
                    break
            if not id_col:
                id_col = df.columns[0]
                
            price_col = None
            for c in ["price", "net price", "unit price", "rate"]:
                for col in df.columns:
                    if str(col).lower() == c:
                        price_col = col
                        break
                if price_col:
                    break
            if not price_col:
                return
                
            date_col = None
            for c in ["invoice date", "datum", "date"]:
                for col in df.columns:
                    if str(col).lower() == c:
                        date_col = col
                        break
                if date_col:
                    break
                    
            df_sorted = df.copy()
            if date_col:
                df_sorted["_parsed_date"] = pd.to_datetime(df_sorted[date_col], dayfirst=True, errors="coerce")
                df_sorted = df_sorted.sort_values(by="_parsed_date", ascending=True)
                
            latest_prices = {}
            for _, r in df_sorted.iterrows():
                sku_val = str(r[id_col]).strip().lower()
                price_val = r[price_col]
                try:
                    price_val = float(price_val)
                    if not pd.isna(price_val) and price_val > 0:
                        latest_prices[sku_val] = price_val
                except (ValueError, TypeError):
                    pass
                    
            enriched_count = 0
            for p in products:
                sku = str(p.get("sku") or "").strip().lower()
                price = p.get("price")
                if (price is None or price == 0.0 or pd.isna(price)) and sku in latest_prices:
                    p["price"] = latest_prices[sku]
                    enriched_count += 1
                    
            print(f"Enriched {enriched_count} product prices from purchase history.")
        except Exception as e:
            print(f"Error enriching prices from history: {e}")

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
