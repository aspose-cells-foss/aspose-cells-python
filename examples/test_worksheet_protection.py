"""
Test Worksheet Protection Feature

This test case verifies the worksheet protection functionality including:
- Setting cell locked state
- Setting formula hidden state
- Protecting worksheet with password
- Verifying protection settings after save/load
"""

import unittest
import os
import sys

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from aspose_cells import Workbook, Style


class TestWorksheetProtection(unittest.TestCase):
    """Test cases for worksheet protection feature."""
    
    def test_worksheet_protection(self):
        """Test worksheet protection with locked cells, hidden formulas, and password protection."""
        print("\n" + "="*60)
        print("Test: Worksheet Protection")
        print("="*60)
        
        # Create a new workbook
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "ProtectedSheet"
        
        # Set up test data
        print("\nSetting up test cells...")
        ws.cells['A1'].value = 1
        ws.cells['A2'].value = 2
        ws.cells['B1'].formula = "=A1+A2"
        ws.cells['C1'].value = "Unlocked cell"
        
        # Set cell C1 to unlocked (locked=False)
        print("  A1: Value = 1 (default locked=True)")
        print("  A2: Value = 2 (default locked=True)")
        print("  B1: Formula = A1+A2 (default locked=True)")
        print("  C1: Value = 'Unlocked cell', locked=False")
        ws.cells['C1'].style.set_locked(False)
        
        # Set B1 formula to hidden
        print("  B1: Formula hidden = True")
        ws.cells['B1'].style.set_formula_hidden(True)
        
        # Protect the worksheet with password
        print("\nProtecting worksheet with password 'abc'...")
        ws.protect(password="abc")
        
        # Verify protection settings
        self.assertTrue(ws.is_protected(), "Worksheet should be protected")
        self.assertEqual(ws.protection['password'], "abc", "Password should be 'abc'")
        self.assertTrue(ws.protection['protected'], "Worksheet protection should be enabled")
        
        # Save the workbook
        output_file = "outputfiles/test_worksheet_protection.xlsx"
        print(f"\nSaving workbook to {output_file}...")
        wb.save(output_file)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_file), f"Output file {output_file} should exist")
        file_size = os.path.getsize(output_file)
        print(f"File size: {file_size} bytes")
        
        # Load the workbook back
        print("\nLoading workbook back to verify protection settings...")
        wb_loaded = Workbook(output_file)
        ws_loaded = wb_loaded.worksheets[0]
        
        # Verify protection settings were persisted
        print("\nVerifying protection settings...")
        self.assertTrue(ws_loaded.name == "ProtectedSheet", "Worksheet name should be preserved")
        self.assertTrue(ws_loaded.is_protected(), "Loaded worksheet should still be protected")
        # Password is stored as legacy hex hash for Excel worksheet protection compatibility
        self.assertIsInstance(ws_loaded.protection['password'], str, "Password should be stored as hex string")
        self.assertEqual(ws_loaded.protection['password'], "CC1A", "Password hash should be CC1A for 'abc'")
        self.assertTrue(ws_loaded.protection['protected'], "Worksheet protection should still be enabled")
        print("\n[OK] Worksheet protection persistence is working correctly!")
        
        # Verify cell values
        print("\nVerifying cell values...")
        self.assertEqual(ws_loaded.cells['A1'].value, 1, "A1 value should be 1")
        self.assertEqual(ws_loaded.cells['A2'].value, 2, "A2 value should be 2")
        self.assertEqual(ws_loaded.cells['B1'].formula, "=A1+A2", "B1 formula should be =A1+A2")
        self.assertEqual(ws_loaded.cells['C1'].value, "Unlocked cell", "C1 value should be 'Unlocked cell'")

        # Verify cell-level protection settings
        print("\nVerifying cell-level protection...")
        self.assertFalse(ws_loaded.cells['C1'].style.protection.locked, "C1 should be unlocked")
        self.assertTrue(ws_loaded.cells['B1'].style.protection.hidden, "B1 formula should be hidden")
        print("  C1: locked=False (unlocked) ✓")
        print("  B1: hidden=True (formula hidden) ✓")

        print("\nAll worksheet protection tests completed successfully!")
        print("="*60 + "\n")
    
    def test_worksheet_unprotect(self):
        """Test unprotecting a worksheet."""
        print("\n" + "="*60)
        print("Test: Worksheet Unprotection")
        print("="*60)
        
        # Create a new workbook
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "UnprotectSheet"
        
        # Protect the worksheet
        print("\nProtecting worksheet...")
        ws.protect(password="test123")
        self.assertTrue(ws.is_protected(), "Worksheet should be protected")
        
        # Unprotect the worksheet
        print("Unprotecting worksheet with password 'test123'...")
        ws.unprotect(password="test123")
        self.assertFalse(ws.is_protected(), "Worksheet should be unprotected")
        self.assertIsNone(ws.protection['password'], "Password should be cleared")
        self.assertFalse(ws.protection['protected'], "Worksheet protection should be disabled")
        
        # Save and verify
        output_file = "outputfiles/test_worksheet_unprotect.xlsx"
        print(f"\nSaving workbook to {output_file}...")
        wb.save(output_file)
        
        self.assertTrue(os.path.exists(output_file), f"Output file {output_file} should exist")
        
        # Load and verify
        print("\nLoading workbook back to verify unprotection...")
        wb_loaded = Workbook(output_file)
        ws_loaded = wb_loaded.worksheets[0]
        
        self.assertFalse(ws_loaded.is_protected(), "Loaded worksheet should be unprotected")
        self.assertIsNone(ws_loaded.protection['password'], "Password should be None")
        
        print("\nAll worksheet unprotection tests completed successfully!")
        print("="*60 + "\n")
    
    def test_worksheet_protection_options(self):
        """Test worksheet protection with various protection options."""
        print("\n" + "="*60)
        print("Test: Worksheet Protection Options")
        print("="*60)
        
        # Create a new workbook
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "ProtectionOptionsSheet"
        
        # Set up test data
        ws.cells['A1'].value = "Test"
        ws.cells['A2'].value = 42
        
        # Protect with specific options
        print("\nProtecting worksheet with custom options...")
        ws.protect(
            password="secure",
            format_cells=True,      # Allow formatting cells
            format_columns=False,    # Don't allow formatting columns
            format_rows=True,        # Allow formatting rows
            insert_columns=False,    # Don't allow inserting columns
            insert_rows=True,        # Allow inserting rows
            delete_columns=False,    # Don't allow deleting columns
            delete_rows=True,        # Allow deleting rows
            sort=True,               # Allow sorting
            auto_filter=True         # Allow auto filter
        )
        
        # Verify protection options
        self.assertTrue(ws.is_protected(), "Worksheet should be protected")
        self.assertEqual(ws.protection['password'], "secure", "Password should be 'secure'")
        self.assertTrue(ws.protection['format_cells'], "format_cells should be True")
        self.assertFalse(ws.protection['format_columns'], "format_columns should be False")
        self.assertTrue(ws.protection['format_rows'], "format_rows should be True")
        self.assertFalse(ws.protection['insert_columns'], "insert_columns should be False")
        self.assertTrue(ws.protection['insert_rows'], "insert_rows should be True")
        self.assertFalse(ws.protection['delete_columns'], "delete_columns should be False")
        self.assertTrue(ws.protection['delete_rows'], "delete_rows should be True")
        self.assertTrue(ws.protection['sort'], "sort should be True")
        self.assertTrue(ws.protection['auto_filter'], "auto_filter should be True")
        
        # Save and verify
        output_file = "outputfiles/test_worksheet_protection_options.xlsx"
        print(f"\nSaving workbook to {output_file}...")
        wb.save(output_file)
        
        self.assertTrue(os.path.exists(output_file), f"Output file {output_file} should exist")
        
        # Load and verify
        print("\nLoading workbook back to verify protection options...")
        wb_loaded = Workbook(output_file)
        ws_loaded = wb_loaded.worksheets[0]
        
        # Verify protection options were persisted
        self.assertTrue(ws_loaded.name == "ProtectionOptionsSheet", "Worksheet name should be preserved")
        self.assertTrue(ws_loaded.is_protected(), "Loaded worksheet should be protected")
        # Password is stored as legacy hex hash for Excel worksheet protection compatibility
        self.assertIsInstance(ws_loaded.protection['password'], str, "Password should be stored as hex string")
        self.assertEqual(ws_loaded.protection['password'], "DC77", "Password hash should be DC77 for 'secure'")
        self.assertTrue(ws_loaded.protection['protected'], "Worksheet protection should still be enabled")
        self.assertTrue(ws_loaded.protection['format_cells'], "format_cells should be preserved")
        self.assertFalse(ws_loaded.protection['format_columns'], "format_columns should be preserved")
        self.assertTrue(ws_loaded.protection['format_rows'], "format_rows should be preserved")
        self.assertFalse(ws_loaded.protection['insert_columns'], "insert_columns should be preserved")
        self.assertTrue(ws_loaded.protection['insert_rows'], "insert_rows should be preserved")
        self.assertFalse(ws_loaded.protection['delete_columns'], "delete_columns should be preserved")
        self.assertTrue(ws_loaded.protection['delete_rows'], "delete_rows should be preserved")
        self.assertTrue(ws_loaded.protection['sort'], "sort should be preserved")
        self.assertTrue(ws_loaded.protection['auto_filter'], "auto_filter should be preserved")
        print("\n[OK] Worksheet protection options persistence is working correctly!")
        
        print("\nAll worksheet protection options tests completed successfully!")
        print("="*60 + "\n")


if __name__ == '__main__':
    # Run the tests
    unittest.main(verbosity=2)
