import sys
import unittest
from pathlib import Path
import pandas as pd
import io
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt

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
            'Art No': ['A100', 'B200.0', 'C300'],  # test float represented as string
            'item': ['Apple Juice', 'Banana Shake', 'Cherry Pie'],
            'Name': ['Golden Apple Juice', 'Yellow Banana Shake', 'Red Cherry Pie'],
            'Category': ['Drinks', 'Drinks', 'Bakery'],
            'Sub-Category': ['Juice', 'Shake', 'Pie'],
            'Steur': [19, 7, 7],
            'unit_price': [1.2, 1.5, 2.5],
            'sale_price': [2.4, 3.0, 5.0],
            'margin_50': [50, 50, 50],
            '7 days': [10, 20, 30],
            'Barcode': ['111111', '222222', '']  # Cherry Pie has empty barcode to test caching
        })
        
        # 2. Create a dummy merge sheet (purchase archives)
        # Note: cherry pie (C300) appears twice to check barcode caching.
        self.df_merge_data = pd.DataFrame({
            'Item No.': ['A100', 'B200', 'C300', 'C300'],
            'Description': ['Apple Juice', 'Banana Shake', 'Cherry Pie', 'Cherry Pie'],
            'Vat%': [19.0, 7.0, 7.0, 7.0],
            'Name': ['Old Apple', 'Old Banana', 'Old Cherry', 'Old Cherry 2']
        })
        
        # Save sheets
        self.loaded_excel_path = self.tmp_dir / "loaded_sheet.xlsx"
        with pd.ExcelWriter(self.loaded_excel_path, engine='openpyxl') as writer:
            self.df_all_data.to_excel(writer, sheet_name="All", index=False)
            
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
                self.tmp_dir.rmdir()
            except OSError:
                pass

    def test_apply_standards_logic(self):
        dialog = SRProductsArchiveDialog(parent=None)
        
        # Simulate loading the sheets_data from Dropbox URL
        dialog.sheets_data = {
            'All': self.df_all_data
        }
        dialog.raw_excel_content = b"fake excel content bytes"
        
        # Set selection paths
        dialog.selected_file_path = str(self.merge_excel_path)
        dialog.selected_file_type = "merge"
        
        # Let's seed the barcode cache beforehand with a barcode for C300 Cherry Pie
        # to test if caching correctly fills it on repeat row
        # In real scenario, the first row of Cherry Pie might have a barcode, or it is cached during comparison.
        # Let's set the first row's barcode in All sheet to test caching during comparison:
        dialog.sheets_data['All'].at[2, 'Barcode'] = '333333' # Chery Pie barcode
        
        # Run apply_standards
        dialog.apply_standards()
        
        # Assert processing succeeded and created 2 sheets
        self.assertIn('purchase archives', dialog.processed_sheets)
        self.assertIn('SR standard Archives', dialog.processed_sheets)
        
        df_processed = dialog.processed_sheets['SR standard Archives']
        
        # Check that Name is replaced
        self.assertEqual(df_processed.at[0, 'Name'], 'Golden Apple Juice')
        self.assertEqual(df_processed.at[1, 'Name'], 'Yellow Banana Shake')
        self.assertEqual(df_processed.at[2, 'Name'], 'Red Cherry Pie')
        self.assertEqual(df_processed.at[3, 'Name'], 'Red Cherry Pie') # Both repeat rows matched description/artno
        
        # Check added columns exist
        cols_to_add = ['Category', 'Sub-Category', 'Steur', 'unit_price', 'sale_price', 'margin_50', '7 days', 'Barcode']
        for col in cols_to_add:
            self.assertIn(col, df_processed.columns)
            
        # Check order: Category should be right after Vat%
        vat_idx = df_processed.columns.get_loc('Vat%')
        cat_idx = df_processed.columns.get_loc('Category')
        self.assertEqual(cat_idx, vat_idx + 1)
        
        # Check values copied
        self.assertEqual(df_processed.at[0, 'Category'], 'Drinks')
        self.assertEqual(df_processed.at[2, 'Category'], 'Bakery')
        
        # Check Barcode caching on repeat row (index 3)
        self.assertEqual(df_processed.at[2, 'Barcode'], '333333')
        self.assertEqual(df_processed.at[3, 'Barcode'], '333333')  # Cached barcode used for repeat row!

if __name__ == "__main__":
    unittest.main()
