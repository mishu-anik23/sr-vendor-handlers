import sys
import unittest
from pathlib import Path
import pandas as pd
from PyQt5.QtWidgets import QApplication

# Ensure path resolution
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from main import SRProductsArchiveDialog

# Initialize QApplication (singleton)
app = QApplication.instance()
if not app:
    app = QApplication(sys.argv)

class TestApplyStandards(unittest.TestCase):
    def setUp(self):
        # Create temp Excel files for testing
        self.tmp_dir = Path(__file__).resolve().parent.parent / "scratch" / "test_data"
        self.tmp_dir.mkdir(parents=True, exist_ok=True)
        
        # 1. Create a dummy 'All' sheet excel
        self.df_all_data = pd.DataFrame({
            'Art No': ['A100', 'B200.0', 'C300'],
            'item': ['Apple Juice', 'Banana Shake', 'Cherry Pie'],
            'Name': ['Golden Apple Juice', 'Yellow Banana Shake', 'Red Cherry Pie'],
            'Category': ['Drinks', 'Drinks', 'Bakery'],
            'Sub-Category': ['Juice', 'Shake', 'Pie'],
            'Steur': [19, 7, 7],
            'unit_price': [1.2, 1.5, 2.5],
            'sale_price': [2.4, 3.0, 5.0],
            'margin_50': [50, 50, 50],
            '7 days': [10, 20, 30],
            'Barcode': ['111111', '222222', ''],
            'Vendor': ['VendorA', 'VendorB', 'VendorC']
        })

        # Create a dummy next sheet ('product') for unique files
        self.df_product_data = pd.DataFrame({
            'Art No': ['A100', 'B200.0', 'C300'],
            'item': ['Apple Juice', 'Banana Shake', 'Cherry Pie'],
            'Name': ['Golden Apple Juice Unique', 'Yellow Banana Shake Unique', 'Red Cherry Pie Unique'],
            'Category': ['Drinks', 'Drinks', 'Bakery'],
            'Sub-Category': ['Juice', 'Shake', 'Pie'],
            'Steur': [19, 7, 7],
            'unit_price': [1.2, 1.5, 2.5],
            'sale_price': [2.4, 3.0, 5.0],
            'margin_50': [50, 50, 50],
            '7 days': [10, 20, 30],
            'Barcode': ['111111', '222222', '333333'],
            'Vendor': ['VendorA', 'VendorB', 'VendorC'],
            'Tag': ['TagA', 'TagB', 'TagC'],
            'Kassen': ['K1', 'K2', 'K3'],
            'Rack': ['R1', 'R2', 'R3']
        })
        
        # 2. Create a dummy merge sheet (purchase archives) with dates out of order to verify sorting
        self.df_merge_data = pd.DataFrame({
            'Invoice Date': ['2026-06-04', '2026-06-01', '2026-06-03', '2026-06-02'],
            'Invoice Number': ['INV-4', 'INV-1', 'INV-3', 'INV-2'],
            'Item No.': ['C300', 'A100', 'C300', 'B200'],
            'Description': ['Cherry Pie', 'Apple Juice', 'Cherry Pie', 'Banana Shake'],
            'Vat%': [7.0, 19.0, 7.0, 7.0],
            'Name': ['Old Cherry 2', 'Old Apple', 'Old Cherry', 'Old Banana']
        })
        
        # Save sheets
        self.loaded_excel_path = self.tmp_dir / "loaded_sheet.xlsx"
        with pd.ExcelWriter(self.loaded_excel_path, engine='openpyxl') as writer:
            self.df_all_data.to_excel(writer, sheet_name="All", index=False)
            self.df_product_data.to_excel(writer, sheet_name="product", index=False)
            
        self.merge_excel_path = self.tmp_dir / "merge_sheet.xlsx"
        with pd.ExcelWriter(self.merge_excel_path, engine='openpyxl') as writer:
            self.df_merge_data.to_excel(writer, sheet_name="Invoice Data", index=False)

    def tearDown(self):
        # Clean up files
        if self.loaded_excel_path.exists():
            self.loaded_excel_path.unlink()
        if self.merge_excel_path.exists():
            self.merge_excel_path.unlink()
        if self.tmp_dir.exists():
            try:
                self.tmp_dir.unlink()
            except Exception:
                pass

    def test_apply_standards_logic_merge(self):
        dialog = SRProductsArchiveDialog(parent=None)
        dialog.selected_vendor_name = "asiaexpress"
        
        # Simulate loading the sheets_data from Dropbox URL
        dialog.sheets_data = {
            'All': self.df_all_data,
            'product': self.df_product_data
        }
        dialog.raw_excel_content = b"fake excel content bytes"
        dialog.selected_file_path = str(self.merge_excel_path)
        dialog.selected_file_type = "merge"
        
        # Seed barcode for Cherry Pie in 'All' sheet
        dialog.sheets_data['All'].at[2, 'Barcode'] = '333333'
        
        # Run apply_sunrise_standard
        dialog.apply_sunrise_standard()
        
        # Assert processing succeeded and created 2 sheets
        self.assertIn('purchase archives', dialog.processed_sheets)
        self.assertIn('SR standard Archives', dialog.processed_sheets)
        
        df_processed = dialog.processed_sheets['SR standard Archives']
        
        # Check order: Row 0 must be INV-1 (Apple Juice, 2026-06-01)
        self.assertEqual(df_processed.at[0, 'Invoice Number'], 'INV-1')
        self.assertEqual(df_processed.at[0, 'Vendor'], 'asian')
        
        # Row 1 must be INV-2 (Banana Shake, 2026-06-02)
        self.assertEqual(df_processed.at[1, 'Invoice Number'], 'INV-2')
        self.assertEqual(df_processed.at[1, 'Vendor'], 'asian1')
        
        # Row 2 must be INV-3 (Cherry Pie, 2026-06-03)
        self.assertEqual(df_processed.at[2, 'Invoice Number'], 'INV-3')
        self.assertEqual(df_processed.at[2, 'Vendor'], 'asian2')
        
        # Row 3 must be INV-4 (Cherry Pie, 2026-06-04)
        self.assertEqual(df_processed.at[3, 'Invoice Number'], 'INV-4')
        self.assertEqual(df_processed.at[3, 'Vendor'], 'asian3')
        
        # Check Barcode caching (Row 2 and Row 3 are Cherry Pie, both should have barcode '333333')
        self.assertEqual(df_processed.at[2, 'Barcode'], '333333')
        self.assertEqual(df_processed.at[3, 'Barcode'], '333333')

    def test_apply_standards_logic_unique(self):
        dialog = SRProductsArchiveDialog(parent=None)
        dialog.selected_vendor_name = "asiaexpress"
        
        # Simulate loading the sheets_data from Dropbox URL
        dialog.sheets_data = {
            'All': self.df_all_data,
            'product': self.df_product_data
        }
        dialog.raw_excel_content = b"fake excel content bytes"
        dialog.selected_file_path = str(self.merge_excel_path)
        dialog.selected_file_type = "unique"
        
        # Run apply_sunrise_standard
        dialog.apply_sunrise_standard()
        
        # Assert processing succeeded and created 2 sheets
        self.assertIn('purchase archives', dialog.processed_sheets)
        self.assertIn('SR standard Archives', dialog.processed_sheets)
        
        df_processed = dialog.processed_sheets['SR standard Archives']
        
        # Deduplication check: INV-4 (Cherry Pie reorder) should have been dropped.
        # Total rows should be 3 (INV-1, INV-2, INV-3)
        self.assertEqual(len(df_processed), 3)
        
        # Apple Juice (INV-1) -> 'asian'
        self.assertEqual(df_processed.at[0, 'Name'], 'Golden Apple Juice Unique')
        self.assertEqual(df_processed.at[0, 'Vendor'], 'asian')
        
        # Banana Shake (INV-2) -> 'asian1'
        self.assertEqual(df_processed.at[1, 'Name'], 'Yellow Banana Shake Unique')
        self.assertEqual(df_processed.at[1, 'Vendor'], 'asian1')
        
        # Cherry Pie (INV-3) -> 'asian2' (since it was the first occurrence)
        self.assertEqual(df_processed.at[2, 'Name'], 'Red Cherry Pie Unique')
        self.assertEqual(df_processed.at[2, 'Vendor'], 'asian2')

if __name__ == "__main__":
    unittest.main()
