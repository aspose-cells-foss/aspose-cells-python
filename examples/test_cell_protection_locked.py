"""
Test Cell Protection (Locked/Hidden) Feature

This test verifies that cell-level protection (locked and hidden properties)
are correctly written to and loaded from XLSX files according to ECMA-376.
"""

import unittest
import os
import sys
import zipfile
import xml.etree.ElementTree as ET

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from aspose_cells import Workbook


class TestCellProtectionLocked(unittest.TestCase):
    """Test cases for cell-level protection (locked/hidden)."""

    def setUp(self):
        """Set up test fixtures."""
        self.ns = {
            'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

    def test_cell_locked_false_roundtrip(self):
        """Test that cells with locked=False can be edited when worksheet is protected."""
        print("\n" + "="*70)
        print("Test: Cell Locked=False Roundtrip")
        print("="*70)

        # Create workbook
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "LockedTest"

        # Set up cells
        print("\nSetting up cells...")
        ws.cells['A1'].value = "Locked (default)"
        ws.cells['A2'].value = "Unlocked"
        ws.cells['A3'].value = "Also unlocked"
        ws.cells['B1'].value = "Hidden formula"
        ws.cells['B1'].formula = "=1+1"

        # Set locked=False for A2 and A3
        print("  A1: Locked (default=True)")
        ws.cells['A2'].style.set_locked(False)
        print("  A2: Locked=False")
        ws.cells['A3'].style.protection.locked = False
        print("  A3: Locked=False (via protection property)")

        # Set formula hidden for B1
        ws.cells['B1'].style.set_formula_hidden(True)
        print("  B1: Formula hidden=True")

        # Protect the worksheet
        ws.protect(password="test123")
        print("\nWorksheet protected with password")

        # Save
        output_file = "outputfiles/test_cell_locked_false.xlsx"
        print(f"\nSaving to {output_file}...")
        wb.save(output_file)
        self.assertTrue(os.path.exists(output_file))

        # Verify XML structure
        print("\nVerifying XML structure...")
        with zipfile.ZipFile(output_file, 'r') as zf:
            # Check styles.xml for protection element
            styles_xml = zf.read('xl/styles.xml').decode('utf-8')
            print(f"Styles XML length: {len(styles_xml)} bytes")

            # Parse and verify protection in cellXfs
            styles_root = ET.fromstring(styles_xml)
            cell_xfs = styles_root.findall('.//main:cellXfs/main:xf', self.ns)
            print(f"Found {len(cell_xfs)} cellXf elements")

            # Find xf with protection element
            found_protection = False
            for xf in cell_xfs:
                prot = xf.find('main:protection', self.ns)
                if prot is not None:
                    locked = prot.get('locked', '1')
                    hidden = prot.get('hidden', '0')
                    print(f"  Found protection: locked={locked}, hidden={hidden}")
                    found_protection = True

            self.assertTrue(found_protection, "Should find at least one protection element in styles")

            # Check worksheet XML for cells with style references
            sheet_xml = zf.read('xl/worksheets/sheet1.xml').decode('utf-8')
            sheet_root = ET.fromstring(sheet_xml)

            # Check that A2 has a style index (s attribute)
            rows = sheet_root.findall('.//main:sheetData/main:row', self.ns)
            a2_style = None
            for row in rows:
                for cell in row.findall('main:c', self.ns):
                    ref = cell.get('r')
                    if ref == 'A2':
                        a2_style = cell.get('s')
                        print(f"  A2 has style index: {a2_style}")

            self.assertIsNotNone(a2_style, "A2 should have a style reference")

        # Load back and verify
        print("\nLoading workbook back...")
        wb_loaded = Workbook(output_file)
        ws_loaded = wb_loaded.worksheets[0]

        # Verify values
        print("Verifying cell values...")
        self.assertEqual(ws_loaded.cells['A1'].value, "Locked (default)")
        self.assertEqual(ws_loaded.cells['A2'].value, "Unlocked")
        self.assertEqual(ws_loaded.cells['A3'].value, "Also unlocked")
        self.assertEqual(ws_loaded.cells['B1'].formula, "=1+1")

        # Verify protection settings
        print("Verifying protection settings...")
        self.assertTrue(ws_loaded.cells['A1'].style.protection.locked, "A1 should be locked")
        self.assertFalse(ws_loaded.cells['A2'].style.protection.locked, "A2 should be unlocked")
        self.assertFalse(ws_loaded.cells['A3'].style.protection.locked, "A3 should be unlocked")
        self.assertTrue(ws_loaded.cells['B1'].style.protection.hidden, "B1 formula should be hidden")

        print("\n[OK] Cell protection roundtrip successful!")
        print("="*70 + "\n")

    def test_protection_in_styles_xml(self):
        """Test that protection element is correctly written to styles.xml."""
        print("\n" + "="*70)
        print("Test: Protection Element in styles.xml")
        print("="*70)

        # Create workbook with various protection settings
        wb = Workbook()
        ws = wb.worksheets[0]

        # Create cells with different protection settings
        ws.cells['A1'].value = "Locked"  # Default: locked=True, hidden=False
        ws.cells['A2'].value = "Unlocked"
        ws.cells['A2'].style.protection.locked = False
        ws.cells['A3'].value = "Hidden formula"
        ws.cells['A3'].formula = "=1+1"
        ws.cells['A3'].style.protection.hidden = True
        ws.cells['A4'].value = "Unlocked and hidden"
        ws.cells['A4'].formula = "=2+2"
        ws.cells['A4'].style.protection.locked = False
        ws.cells['A4'].style.protection.hidden = True

        # Save
        output_file = "outputfiles/test_protection_styles.xlsx"
        print(f"\nSaving to {output_file}...")
        wb.save(output_file)

        # Verify styles.xml
        print("\nVerifying styles.xml structure...")
        with zipfile.ZipFile(output_file, 'r') as zf:
            styles_xml = zf.read('xl/styles.xml').decode('utf-8')
            styles_root = ET.fromstring(styles_xml)

            cell_xfs = styles_root.findall('.//main:cellXfs/main:xf', self.ns)
            print(f"Found {len(cell_xfs)} cellXf elements")

            protection_count = 0
            for i, xf in enumerate(cell_xfs):
                prot = xf.find('main:protection', self.ns)
                if prot is not None:
                    locked = prot.get('locked', '1')
                    hidden = prot.get('hidden', '0')
                    print(f"  cellXf[{i}]: locked={locked}, hidden={hidden}")
                    protection_count += 1

            # We should have at least one protection element
            self.assertGreater(protection_count, 0, "Should have protection elements in styles")
            print(f"\n[OK] Found {protection_count} protection elements")

        print("="*70 + "\n")

    def test_unlocked_cell_with_sheet_protection(self):
        """Test complete scenario: unlocked cells should be editable when sheet is protected."""
        print("\n" + "="*70)
        print("Test: Unlocked Cells with Sheet Protection")
        print("="*70)

        # Create workbook
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "EditableWhenProtected"

        # Set up a simple data entry form
        print("\nCreating data entry form...")
        ws.cells['A1'].value = "Name:"
        ws.cells['B1'].value = ""  # Editable field
        ws.cells['B1'].style.set_locked(False)

        ws.cells['A2'].value = "Age:"
        ws.cells['B2'].value = ""  # Editable field
        ws.cells['B2'].style.set_locked(False)

        ws.cells['A3'].value = "Total:"
        ws.cells['B3'].formula = "=B2*2"  # Locked, calculated field

        print("  A1, A2, A3: Labels (locked)")
        print("  B1, B2: Input fields (unlocked)")
        print("  B3: Calculated field (locked)")

        # Protect the worksheet
        ws.protect(password="form123")
        print("\nWorksheet protected")

        # Save and reload
        output_file = "outputfiles/test_unlocked_cells_protected.xlsx"
        print(f"\nSaving to {output_file}...")
        wb.save(output_file)

        print("Loading back...")
        wb_loaded = Workbook(output_file)
        ws_loaded = wb_loaded.worksheets[0]

        # Verify protection states
        print("\nVerifying protection states...")
        self.assertTrue(ws_loaded.is_protected(), "Worksheet should be protected")
        self.assertTrue(ws_loaded.cells['A1'].style.protection.locked, "A1 (label) should be locked")
        self.assertFalse(ws_loaded.cells['B1'].style.protection.locked, "B1 (input) should be unlocked")
        self.assertFalse(ws_loaded.cells['B2'].style.protection.locked, "B2 (input) should be unlocked")
        self.assertTrue(ws_loaded.cells['B3'].style.protection.locked, "B3 (formula) should be locked")

        print("  [OK] A1: locked")
        print("  [OK] B1: unlocked")
        print("  [OK] B2: unlocked")
        print("  [OK] B3: locked")

        print("\n[OK] Unlocked cells correctly identified!")
        print("="*70 + "\n")


if __name__ == '__main__':
    # Create output directory if it doesn't exist
    os.makedirs('outputfiles', exist_ok=True)

    # Run tests
    unittest.main(verbosity=2)
