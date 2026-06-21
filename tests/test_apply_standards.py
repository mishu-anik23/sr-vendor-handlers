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
        
        # 2. Create a dummy merge sheet (purchase archives)
        self.df_merge_data = pd.DataFrame({
            'Invoice Date': ['2026-06-01', '2026-06-02', '2026-06-03', '2026-06-04'],
            'Invoice Number': ['INV-1', 'INV-2', 'INV-3', 'INV-4'],
            'Item No.': ['A100', 'B200', 'C300', 'C300'],
            'Description': ['Apple Juice', 'Banana Shake', 'Cherry Pie', 'Cherry Pie'],
            'Vat%': [19.0, 7.0, 7.0, 7.0],
            'Name': ['Old Apple', 'Old Banana', 'Old Cherry', 'Old Cherry 2'],
            'Vendor': ['VendorA', 'VendorB', 'VendorC', 'VendorC']
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
        
        # Simulate loading the sheets_data from Dropbox URL
        dialog.sheets_data = {
            'All': self.df_all_data,
            'product': self.df_product_data
        }
        dialog.raw_excel_content = b"fake excel content bytes"
        dialog.selected_file_path = str(self.merge_excel_path)
        dialog.selected_file_type = "merge"
        
        # Seed barcode
        dialog.sheets_data['All'].at[2, 'Barcode'] = '333333'
        
        # Run apply_sunrise_standard
        dialog.apply_sunrise_standard()
        
        # Assert processing succeeded and created 2 sheets
        self.assertIn('purchase archives', dialog.processed_sheets)
        self.assertIn('SR standard Archives', dialog.processed_sheets)
        
        df_processed = dialog.processed_sheets['SR standard Archives']
        
        # Check that Name is replaced
        self.assertEqual(df_processed.at[0, 'Name'], 'Golden Apple Juice')
        
        # Check added columns exist
        cols_to_add = ['Category', 'Sub-Category', 'Steur', 'unit_price', 'sale_price', 'margin_50', '7 days', 'Barcode', 'Vendor']
        for col in cols_to_add:
            self.assertIn(col, df_processed.columns)
            
        # Check order: Vendor should be right between 'Invoice Date' and 'Invoice Number'
        date_idx = df_processed.columns.get_loc('Invoice Date')
        vendor_idx = df_processed.columns.get_loc('Vendor')
        num_idx = df_processed.columns.get_loc('Invoice Number')
        self.assertEqual(vendor_idx, date_idx + 1)
        self.assertEqual(num_idx, vendor_idx + 1)
        
        # Check Vendor value copied correctly
        self.assertEqual(df_processed.at[0, 'Vendor'], 'VendorA')
        self.assertEqual(df_processed.at[2, 'Vendor'], 'VendorC')

    def test_apply_standards_logic_unique(self):
        dialog = SRProductsArchiveDialog(parent=None)
        
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
        
        # Check that product sheet values were loaded (e.g. contains 'Unique' suffix)
        self.assertEqual(df_processed.at[0, 'Name'], 'Golden Apple Juice Unique')
        
        # Check added columns exist, including Tag, Kassen, Rack
        cols_to_add = ['Category', 'Sub-Category', 'Steur', 'unit_price', 'sale_price', 'margin_50', '7 days', 'Barcode', 'Vendor', 'Tag', 'Kassen', 'Rack']
        for col in cols_to_add:
            self.assertIn(col, df_processed.columns)
            
        # Check order: Tag, Kassen, Rack should be inserted right after Steur
        steur_idx = df_processed.columns.get_loc('Steur')
        tag_idx = df_processed.columns.get_loc('Tag')
        kassen_idx = df_processed.columns.get_loc('Kassen')
        rack_idx = df_processed.columns.get_loc('Rack')
        
        self.assertEqual(tag_idx, steur_idx + 1)
        self.assertEqual(kassen_idx, tag_idx + 1)
        self.assertEqual(rack_idx, kassen_idx + 1)
        
        # Check Tag, Kassen, Rack values copied correctly
        self.assertEqual(df_processed.at[0, 'Tag'], 'TagA')
        self.assertEqual(df_processed.at[1, 'Kassen'], 'K2')
        self.assertEqual(df_processed.at[2, 'Rack'], 'R3')

if __name__ == "__main__":
    unittest.main()
