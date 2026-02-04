"""
Test workbook protection functionality.

This test verifies that:
1. A workbook can be created with data
2. All worksheets can be protected with a password
3. The protected workbook can be saved
4. The workbook can be loaded and data verified
"""

import unittest
import os
import sys

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from aspose_cells import Workbook


class TestWorkbookProtection(unittest.TestCase):
    """Test cases for workbook protection functionality."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.output_dir = os.path.join(os.path.dirname(__file__), '..', 'outputfiles')
        os.makedirs(self.output_dir, exist_ok=True)
        self.test_file = os.path.join(self.output_dir, 'test_workbook_protection.xlsx')
        self.password = "TestPassword123"
        
        # Clean up any existing test file
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
    
    def tearDown(self):
        """Clean up after tests."""
        # Optionally keep the file for manual inspection
        pass
    
    def test_workbook_protection(self):
        """
        Test workbook protection:
        1. Create a workbook and populate with data
        2. Protect all worksheets with a password
        3. Save the workbook
        4. Load the workbook and verify all data is correct
        """
        print("\n" + "="*70)
        print("TEST: Workbook Protection")
        print("="*70)
        
        # Step 1: Create a workbook and populate with data
        print("\nStep 1: Creating workbook and populating with data...")
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "TestData"
        
        # Populate with various types of data
        test_data = {
            'A1': "Name",
            'B1': "Age",
            'C1': "Score",
            'A2': "Alice",
            'B2': 25,
            'C2': 95.5,
            'A3': "Bob",
            'B3': 30,
            'C3': 87.3,
            'A4': "Charlie",
            'B4': 28,
            'C4': 92.1,
            'A5': "Total",
            'B5': 83,  # Sum of ages
            'C5': 274.9  # Sum of scores
        }
        
        for cell_ref, value in test_data.items():
            ws.cells[cell_ref].value = value
            print(f"  Set {cell_ref} = {value}")
        
        # Add some styling
        ws.cells['A1'].style.font.bold = True
        ws.cells['B1'].style.font.bold = True
        ws.cells['C1'].style.font.bold = True
        
        # Store original data for verification
        original_data = {}
        for ref, cell in ws.cells._cells.items():
            original_data[ref] = cell.value
        
        print(f"  Total cells populated: {len(original_data)}")
        
        # Step 2: Protect all worksheets with a password
        print(f"\nStep 2: Protecting worksheet with password...")
        ws.protect(self.password)
        print(f"  Worksheet protected with password: '{self.password}'")
        print(f"  Worksheet is protected: {ws.is_protected()}")
        print(f"  Protection password set: {ws.protection.get('password')}")
        
        # Step 3: Save the workbook
        print(f"\nStep 3: Saving workbook to file...")
        wb.save(self.test_file)
        file_size = os.path.getsize(self.test_file)
        print(f"  Saved to: {self.test_file}")
        print(f"  File size: {file_size} bytes")
        
        # Step 4: Load the workbook and verify all data
        print(f"\nStep 4: Loading workbook and verifying data...")
        wb_loaded = Workbook(self.test_file)
        ws_loaded = wb_loaded.worksheets[0]
        
        print(f"  Loaded worksheet name: {ws_loaded.name}")
        print(f"  Worksheet is protected: {ws_loaded.is_protected()}")
        print(f"  Protection password: {ws_loaded.protection.get('password')}")
        
        # Verify all data
        print(f"\nVerifying cell data...")
        all_correct = True
        mismatched_cells = []
        
        for ref, expected_value in original_data.items():
            actual_value = ws_loaded.cells[ref].value
            
            # Handle floating point comparison
            if isinstance(expected_value, float) and isinstance(actual_value, float):
                is_correct = abs(expected_value - actual_value) < 0.001
            else:
                is_correct = expected_value == actual_value
            
            if is_correct:
                print(f"  [OK] {ref}: {actual_value} (correct)")
            else:
                print(f"  [FAIL] {ref}: expected {expected_value}, got {actual_value} (MISMATCH)")
                all_correct = False
                mismatched_cells.append(ref)
        
        # Verify protection status
        print(f"\nVerifying protection status...")
        # Note: password is hashed, so we verify it's set (not None) and protection is enabled
        protection_correct = ws_loaded.is_protected() and ws_loaded.protection.get('password') is not None
        if protection_correct:
            print(f"  [OK] Protection status correct (password hash: '{ws_loaded.protection.get('password')}')")
        else:
            print(f"  [FAIL] Protection status incorrect")
            print(f"    Expected: protected=True, password='{self.password}'")
            print(f"    Got: protected={ws_loaded.is_protected()}, password='{ws_loaded.protection.get('password')}'")
            all_correct = False
        
        # Final result
        print("\n" + "="*70)
        if all_correct:
            print("RESULT: [PASS] ALL TESTS PASSED")
            print("="*70)
            print(f"  - {len(original_data)} cells verified")
            print(f"  - Protection status verified")
            print(f"  - Data integrity confirmed")
        else:
            print("RESULT: [FAIL] TESTS FAILED")
            print("="*70)
            if mismatched_cells:
                print(f"  - {len(mismatched_cells)} cells with mismatched data")
            if not protection_correct:
                print(f"  - Protection status incorrect")
        print("="*70)
        
        self.assertTrue(all_correct, "Workbook protection test failed")
    
    def test_multiple_worksheets_protection(self):
        """
        Test protecting multiple worksheets in a workbook.
        """
        print("\n" + "="*70)
        print("TEST: Multiple Worksheets Protection")
        print("="*70)
        
        # Create workbook with multiple worksheets
        wb = Workbook()
        wb.worksheets[0].name = "Sheet1"
        ws2 = wb.create_worksheet("Sheet2")
        ws3 = wb.create_worksheet("Sheet3")
        
        # Populate data in all sheets
        for i, ws in enumerate(wb.worksheets, 1):
            ws.cells[f'A{i}'].value = f"Sheet{i} Data"
            ws.cells[f'B{i}'].value = i * 100
            print(f"  Populated {ws.name}: A{i}='{ws.cells[f'A{i}'].value}', B{i}={ws.cells[f'B{i}'].value}")
        
        # Protect all worksheets with the same password
        print(f"\nProtecting all worksheets with password...")
        for ws in wb.worksheets:
            ws.protect(self.password)
            print(f"  - {ws.name}: protected")
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_multi_sheet_protection.xlsx')
        wb.save(test_file)
        print(f"\nSaved to: {test_file}")
        
        # Load and verify
        wb_loaded = Workbook(test_file)
        print(f"\nLoaded workbook with {len(wb_loaded.worksheets)} worksheets")
        
        all_correct = True
        for i, ws in enumerate(wb_loaded.worksheets, 1):
            expected_name = f"Sheet{i}"
            expected_a = f"Sheet{i} Data"
            expected_b = i * 100
            
            name_correct = ws.name == expected_name
            a_correct = ws.cells[f'A{i}'].value == expected_a
            b_correct = ws.cells[f'B{i}'].value == expected_b
            protected = ws.is_protected()
            
            if name_correct and a_correct and b_correct and protected:
                print(f"  [OK] {ws.name}: data correct, protected")
            else:
                print(f"  [FAIL] {ws.name}: FAILED")
                if not name_correct:
                    print(f"    - Name: expected '{expected_name}', got '{ws.name}'")
                if not a_correct:
                    print(f"    - A{i}: expected '{expected_a}', got '{ws.cells[f'A{i}'].value}'")
                if not b_correct:
                    print(f"    - B{i}: expected {expected_b}, got {ws.cells[f'B{i}'].value}")
                if not protected:
                    print(f"    - Protection: expected True, got False")
                all_correct = False
        
        print("\n" + "="*70)
        if all_correct:
            print("RESULT: [PASS] ALL TESTS PASSED")
        else:
            print("RESULT: [FAIL] TESTS FAILED")
        print("="*70)
        
        self.assertTrue(all_correct, "Multiple worksheets protection test failed")
    
    def test_protection_without_password(self):
        """
        Test worksheet protection without a password.
        """
        print("\n" + "="*70)
        print("TEST: Protection Without Password")
        print("="*70)
        
        # Create and populate workbook
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = "Test Data"
        ws.cells['B1'].value = 123
        
        # Protect without password
        print("\nProtecting worksheet without password...")
        ws.protect()
        print(f"  Worksheet protected: {ws.is_protected()}")
        print(f"  Password: {ws.protection.get('password')}")
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_protection_no_password.xlsx')
        wb.save(test_file)
        print(f"\nSaved to: {test_file}")
        
        # Load and verify
        wb_loaded = Workbook(test_file)
        ws_loaded = wb_loaded.worksheets[0]
        
        data_correct = (
            ws_loaded.cells['A1'].value == "Test Data" and
            ws_loaded.cells['B1'].value == 123
        )
        protected_correct = ws_loaded.is_protected()
        password_correct = ws_loaded.protection.get('password') is None
        
        print(f"\nVerification:")
        print(f"  Data correct: {data_correct}")
        print(f"  Protected: {protected_correct}")
        print(f"  No password: {password_correct}")
        
        all_correct = data_correct and protected_correct and password_correct
        
        print("\n" + "="*70)
        if all_correct:
            print("RESULT: [PASS] ALL TESTS PASSED")
        else:
            print("RESULT: [FAIL] TESTS FAILED")
        print("="*70)
        
        self.assertTrue(all_correct, "Protection without password test failed")


if __name__ == '__main__':
    # Run tests with verbose output
    unittest.main(verbosity=2)
