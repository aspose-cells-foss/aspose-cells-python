import unittest
import os
import sys

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose_cells import Workbook, Worksheet, Cell


class TestWorksheetManagement(unittest.TestCase):
    """
    Test comprehensive worksheet management functionality with save/load verification.
    
    Features tested:
    - Create worksheets (with default name, with custom name)
    - Delete worksheets (by index, by name, by Worksheet object)
    - Rename worksheets (using name property, using rename method)
    - Copy worksheets (by index, by name, by Worksheet object)
    - Access worksheets by name
    - Access worksheets by index
    - Get active worksheet
    - Set active worksheet (by index, by name, by Worksheet object)
    - Comprehensive worksheet management (create, delete, rename, copy, access, set active)
    - Edge cases (delete only worksheet, copy non-existent, invalid index, invalid active, duplicate names)
    - Worksheet management API methods
    """
    
    def setUp(self):
        """Set up test workbook."""
        self.workbook = Workbook()
    
    def test_create_worksheets(self):
        """Test creating worksheets."""
        # Test creating worksheets with default name
        ws1 = self.workbook.create_worksheet()
        self.assertIsNotNone(ws1)
        self.assertEqual(ws1.name, 'Sheet2')  # First default is Sheet1
        
        ws2 = self.workbook.create_worksheet()
        self.assertIsNotNone(ws2)
        self.assertEqual(ws2.name, 'Sheet3')
        
        # Test creating worksheets with custom name
        ws3 = self.workbook.create_worksheet("CustomSheet1")
        self.assertIsNotNone(ws3)
        self.assertEqual(ws3.name, 'CustomSheet1')
        
        ws4 = self.workbook.create_worksheet("CustomSheet2")
        self.assertIsNotNone(ws4)
        self.assertEqual(ws4.name, 'CustomSheet2')
        
        # Verify total number of worksheets
        self.assertEqual(len(self.workbook.worksheets), 5)
        
        print("Created 5 worksheets successfully")
    
    def test_delete_worksheets(self):
        """Test deleting worksheets."""
        # Create multiple worksheets
        ws1 = self.workbook.create_worksheet("SheetToDelete1")
        ws2 = self.workbook.create_worksheet("SheetToDelete2")
        ws3 = self.workbook.create_worksheet("SheetToDelete3")
        
        initial_count = len(self.workbook.worksheets)
        self.assertEqual(initial_count, 4)  # Sheet1 + 3 new sheets
        
        # Delete by index
        self.workbook.remove_worksheet(1)  # Delete SheetToDelete1
        self.assertEqual(len(self.workbook.worksheets), 3)
        
        # Delete by name
        self.workbook.remove_worksheet("SheetToDelete2")
        self.assertEqual(len(self.workbook.worksheets), 2)
        
        # Delete by Worksheet object
        self.workbook.remove_worksheet(ws3)
        self.assertEqual(len(self.workbook.worksheets), 1)
        
        print("Deleted 3 worksheets successfully")
    
    def test_rename_worksheets(self):
        """Test renaming worksheets."""
        # Create worksheets
        ws1 = self.workbook.create_worksheet("OriginalName1")
        ws2 = self.workbook.create_worksheet("OriginalName2")
        
        # Rename using name property
        ws1.name = "RenamedName1"
        self.assertEqual(ws1.name, "RenamedName1")
        
        # Rename using rename method
        ws2.rename("RenamedName2")
        self.assertEqual(ws2.name, "RenamedName2")
        
        print("Renamed 2 worksheets successfully")
    
    def test_copy_worksheets(self):
        """Test copying worksheets."""
        # Create a worksheet with some data
        ws1 = self.workbook.create_worksheet("OriginalSheet")
        ws1.cells["A1"] = Cell("Original Data")
        ws1.cells["A2"] = Cell("Row 2")
        ws1.cells["B1"] = Cell("Column B")
        ws1.cells["B2"] = Cell("Data B2")
        
        # Copy by index
        ws_copy1 = self.workbook.copy_worksheet(1)  # Index of OriginalSheet
        self.assertIsNotNone(ws_copy1)
        self.assertEqual(ws_copy1.name, "OriginalSheet (copy)")
        
        # Verify data was copied
        self.assertEqual(ws_copy1.cells["A1"].value, "Original Data")
        self.assertEqual(ws_copy1.cells["A2"].value, "Row 2")
        self.assertEqual(ws_copy1.cells["B1"].value, "Column B")
        self.assertEqual(ws_copy1.cells["B2"].value, "Data B2")
        
        # Copy by name
        ws_copy2 = self.workbook.copy_worksheet("OriginalSheet")
        self.assertIsNotNone(ws_copy2)
        self.assertEqual(ws_copy2.name, "OriginalSheet (copy1)")
        
        # Copy by Worksheet object
        ws_copy3 = self.workbook.copy_worksheet(ws1)
        self.assertIsNotNone(ws_copy3)
        self.assertEqual(ws_copy3.name, "OriginalSheet (copy2)")
        
        # Verify total number of worksheets (Sheet1 + OriginalSheet + 3 copies = 5)
        self.assertEqual(len(self.workbook.worksheets), 5)
        
        print("Copied worksheet 3 times successfully")
    
    def test_access_worksheets_by_name(self):
        """Test accessing worksheets by name."""
        # Create worksheets with specific names
        ws1 = self.workbook.create_worksheet("FirstSheet")
        ws2 = self.workbook.create_worksheet("SecondSheet")
        ws3 = self.workbook.create_worksheet("ThirdSheet")
        
        # Access by name
        found_ws1 = self.workbook.get_worksheet_by_name("FirstSheet")
        self.assertIsNotNone(found_ws1)
        self.assertEqual(found_ws1.name, "FirstSheet")
        self.assertIs(found_ws1, ws1)
        
        found_ws2 = self.workbook.get_worksheet_by_name("SecondSheet")
        self.assertIsNotNone(found_ws2)
        self.assertEqual(found_ws2.name, "SecondSheet")
        self.assertIs(found_ws2, ws2)
        
        found_ws3 = self.workbook.get_worksheet_by_name("ThirdSheet")
        self.assertIsNotNone(found_ws3)
        self.assertEqual(found_ws3.name, "ThirdSheet")
        self.assertIs(found_ws3, ws3)
        
        # Test non-existent worksheet
        not_found = self.workbook.get_worksheet_by_name("NonExistent")
        self.assertIsNone(not_found)
        
        print("Accessed worksheets by name successfully")
    
    def test_access_worksheets_by_index(self):
        """Test accessing worksheets by index."""
        # Create worksheets
        ws1 = self.workbook.create_worksheet("IndexSheet1")
        ws2 = self.workbook.create_worksheet("IndexSheet2")
        ws3 = self.workbook.create_worksheet("IndexSheet3")
        
        # Access by index
        found_ws0 = self.workbook.get_worksheet_by_index(0)
        self.assertIsNotNone(found_ws0)
        self.assertEqual(found_ws0.name, "Sheet1")
        
        found_ws1 = self.workbook.get_worksheet_by_index(1)
        self.assertIsNotNone(found_ws1)
        self.assertEqual(found_ws1.name, "IndexSheet1")
        self.assertIs(found_ws1, ws1)
        
        found_ws2 = self.workbook.get_worksheet_by_index(2)
        self.assertIsNotNone(found_ws2)
        self.assertEqual(found_ws2.name, "IndexSheet2")
        self.assertIs(found_ws2, ws2)
        
        found_ws3 = self.workbook.get_worksheet_by_index(3)
        self.assertIsNotNone(found_ws3)
        self.assertEqual(found_ws3.name, "IndexSheet3")
        self.assertIs(found_ws3, ws3)
        
        # Test invalid index
        not_found = self.workbook.get_worksheet_by_index(100)
        self.assertIsNone(not_found)
        
        not_found2 = self.workbook.get_worksheet_by_index(-1)
        self.assertIsNone(not_found2)
        
        print("Accessed worksheets by index successfully")
    
    def test_get_active_worksheet(self):
        """Test getting the active worksheet."""
        # Default active worksheet should be the first one
        active_ws = self.workbook.get_active_worksheet()
        self.assertIsNotNone(active_ws)
        self.assertEqual(active_ws.name, "Sheet1")
        self.assertEqual(active_ws, self.workbook.worksheets[0])
        
        # Create more worksheets
        ws2 = self.workbook.create_worksheet("SecondSheet")
        ws3 = self.workbook.create_worksheet("ThirdSheet")
        
        # Active worksheet should still be the first one
        active_ws = self.workbook.get_active_worksheet()
        self.assertEqual(active_ws.name, "Sheet1")
        
        print("Got active worksheet successfully")
    
    def test_set_active_worksheet(self):
        """Test setting the active worksheet."""
        # Create worksheets
        ws1 = self.workbook.worksheets[0]
        ws2 = self.workbook.create_worksheet("ActiveSheet1")
        ws3 = self.workbook.create_worksheet("ActiveSheet2")
        
        # Set active by index
        self.workbook.set_active_worksheet(1)
        active_ws = self.workbook.get_active_worksheet()
        self.assertEqual(active_ws.name, "ActiveSheet1")
        
        # Set active by name
        self.workbook.set_active_worksheet("ActiveSheet2")
        active_ws = self.workbook.get_active_worksheet()
        self.assertEqual(active_ws.name, "ActiveSheet2")
        
        # Set active by Worksheet object
        self.workbook.set_active_worksheet(ws1)
        active_ws = self.workbook.get_active_worksheet()
        self.assertEqual(active_ws.name, "Sheet1")
        
        print("Set active worksheet successfully")
    
    def test_comprehensive_worksheet_management(self):
        """Test all worksheet management features and save to Excel file."""
        print("Testing comprehensive worksheet management...")
        
        # 1. Create multiple worksheets
        print("  Creating worksheets...")
        ws1 = self.workbook.worksheets[0]
        ws1.name = "MainSheet"
        ws1.cells["A1"] = Cell("Main Worksheet")
        ws1.cells["A2"] = Cell("This is the primary worksheet")
        
        ws2 = self.workbook.create_worksheet("DataSheet")
        ws2.cells["A1"] = Cell("Data Worksheet")
        ws2.cells["A2"] = Cell("Contains data tables")
        
        ws3 = self.workbook.create_worksheet("ReportSheet")
        ws3.cells["A1"] = Cell("Report Worksheet")
        ws3.cells["A2"] = Cell("Contains reports")
        
        ws4 = self.workbook.create_worksheet("ConfigSheet")
        ws4.cells["A1"] = Cell("Config Worksheet")
        ws4.cells["A2"] = Cell("Contains configuration")
        
        ws5 = self.workbook.create_worksheet("TempSheet")
        ws5.cells["A1"] = Cell("Temporary Worksheet")
        ws5.cells["A2"] = Cell("Will be deleted")
        
        print(f"    Created {len(self.workbook.worksheets)} worksheets")
        
        # 2. Set active worksheet
        print("  Setting active worksheet...")
        self.workbook.set_active_worksheet("DataSheet")
        active_ws = self.workbook.get_active_worksheet()
        self.assertEqual(active_ws.name, "DataSheet")
        print(f"    Active worksheet: {active_ws.name}")
        
        # 3. Copy a worksheet
        print("  Copying worksheet...")
        ws_copy = self.workbook.copy_worksheet("DataSheet")
        self.assertIsNotNone(ws_copy)
        self.assertEqual(ws_copy.name, "DataSheet (copy)")
        ws_copy.cells["A3"] = Cell("This is a copy")
        print(f"    Copied worksheet: {ws_copy.name}")
        
        # 4. Rename a worksheet
        print("  Renaming worksheet...")
        ws5.name = "RenamedSheet"
        ws5.cells["A1"] = Cell("Renamed Worksheet")
        ws5.cells["A2"] = Cell("This sheet was renamed")
        print(f"    Renamed worksheet to: {ws5.name}")
        
        # 5. Access worksheets by name
        print("  Accessing worksheets by name...")
        found_ws = self.workbook.get_worksheet_by_name("MainSheet")
        self.assertIsNotNone(found_ws)
        self.assertEqual(found_ws.name, "MainSheet")
        print(f"    Found worksheet by name: {found_ws.name}")
        
        # 6. Access worksheets by index
        print("  Accessing worksheets by index...")
        found_ws = self.workbook.get_worksheet_by_index(2)
        self.assertIsNotNone(found_ws)
        self.assertEqual(found_ws.name, "ReportSheet")
        print(f"    Found worksheet by index 2: {found_ws.name}")
        
        # 7. Delete a worksheet (note: TempSheet was renamed to RenamedSheet, so we delete RenamedSheet)
        print("  Deleting worksheet...")
        initial_count = len(self.workbook.worksheets)
        self.workbook.remove_worksheet("RenamedSheet")
        new_count = len(self.workbook.worksheets)
        self.assertEqual(new_count, initial_count - 1)
        print(f"    Deleted worksheet. Count: {initial_count} -> {new_count}")
        
        # 8. Verify all remaining worksheets
        print("  Verifying remaining worksheets...")
        expected_names = ["MainSheet", "DataSheet", "ReportSheet", "ConfigSheet", "DataSheet (copy)"]
        actual_names = [ws.name for ws in self.workbook.worksheets]
        for expected_name in expected_names:
            self.assertIn(expected_name, actual_names)
        print(f"    Verified {len(expected_names)} worksheets")
        
        # 9. Set another active worksheet
        print("  Setting new active worksheet...")
        self.workbook.set_active_worksheet(0)
        active_ws = self.workbook.get_active_worksheet()
        self.assertEqual(active_ws.name, "MainSheet")
        print(f"    New active worksheet: {active_ws.name}")
        
        # Save workbook to outputfiles folder
        output_path = 'outputfiles/test_worksheet_management.xlsx'
        
        # Ensure outputfiles directory exists
        os.makedirs('outputfiles', exist_ok=True)
        
        print(f"Saving workbook to {output_path}...")
        self.workbook.save(output_path)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify file is not empty
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"Worksheet management test file saved to: {output_path}")
        print(f"File size: {file_size} bytes")
        print(f"Total worksheets: {len(self.workbook.worksheets)}")
        
        return {
            'worksheet_count': len(self.workbook.worksheets),
            'active_worksheet': active_ws.name,
            'worksheet_names': actual_names
        }
    
    def test_verify_worksheet_management(self):
        """Test reading generated files and verify all worksheet management settings are correct."""
        # First create comprehensive worksheet management
        original_data = self.test_comprehensive_worksheet_management()
        
        # Load the file back and verify settings
        print("Loading file back and verifying worksheet management settings...")
        loaded_workbook = Workbook('outputfiles/test_worksheet_management.xlsx')
        
        # Verify total number of worksheets
        print("Verifying worksheet count...")
        self.assertEqual(len(loaded_workbook.worksheets), original_data['worksheet_count'])
        print(f"  Worksheet count: {len(loaded_workbook.worksheets)}")
        
        # Verify all worksheet names exist
        print("Verifying worksheet names...")
        for expected_name in original_data['worksheet_names']:
            found_ws = loaded_workbook.get_worksheet_by_name(expected_name)
            self.assertIsNotNone(found_ws, f"Worksheet '{expected_name}' not found")
            print(f"  Found: {expected_name}")
        
        # Verify active worksheet
        print("Verifying active worksheet...")
        active_ws = loaded_workbook.get_active_worksheet()
        self.assertIsNotNone(active_ws)
        self.assertEqual(active_ws.name, original_data['active_worksheet'])
        print(f"  Active worksheet: {active_ws.name}")
        
        # Verify worksheet data
        print("Verifying worksheet data...")
        main_ws = loaded_workbook.get_worksheet_by_name("MainSheet")
        self.assertIsNotNone(main_ws)
        self.assertEqual(main_ws.cells["A1"].value, "Main Worksheet")
        self.assertEqual(main_ws.cells["A2"].value, "This is the primary worksheet")
        print(f"  MainSheet data verified")
        
        data_ws = loaded_workbook.get_worksheet_by_name("DataSheet")
        self.assertIsNotNone(data_ws)
        self.assertEqual(data_ws.cells["A1"].value, "Data Worksheet")
        self.assertEqual(data_ws.cells["A2"].value, "Contains data tables")
        print(f"  DataSheet data verified")
        
        copy_ws = loaded_workbook.get_worksheet_by_name("DataSheet (copy)")
        self.assertIsNotNone(copy_ws)
        self.assertEqual(copy_ws.cells["A1"].value, "Data Worksheet")
        self.assertEqual(copy_ws.cells["A2"].value, "Contains data tables")
        self.assertEqual(copy_ws.cells["A3"].value, "This is a copy")
        print(f"  DataSheet (copy) data verified")
        
        # Note: RenamedSheet was deleted in the comprehensive test
        
        print("All worksheet management settings verified successfully!")
    
    def test_worksheet_management_edge_cases(self):
        """Test edge cases for worksheet management."""
        # Test deleting the only worksheet (should leave at least one)
        wb = Workbook()
        initial_count = len(wb.worksheets)
        self.assertEqual(initial_count, 1)
        
        # Try to delete the only worksheet (implementation may prevent this)
        try:
            wb.remove_worksheet(0)
            # If it succeeds, check that there's still at least one worksheet
            self.assertGreaterEqual(len(wb.worksheets), 1)
        except Exception as e:
            # If it raises an exception, that's also acceptable
            self.assertIsNotNone(e)
        
        # Test copying a non-existent worksheet
        wb = Workbook()
        result = wb.copy_worksheet("NonExistent")
        self.assertIsNone(result)
        
        # Test accessing worksheet with invalid index
        wb = Workbook()
        result = wb.get_worksheet_by_index(-1)
        self.assertIsNone(result)
        
        result = wb.get_worksheet_by_index(100)
        self.assertIsNone(result)
        
        # Test setting active worksheet with invalid index
        wb = Workbook()
        # This should either do nothing or raise an exception
        try:
            wb.set_active_worksheet(100)
            # If it succeeds, active worksheet should still be valid
            active = wb.get_active_worksheet()
            self.assertIsNotNone(active)
        except Exception as e:
            # If it raises an exception, that's also acceptable
            self.assertIsNotNone(e)
        
        # Test creating worksheet with duplicate name
        wb = Workbook()
        ws1 = wb.create_worksheet("Duplicate")
        ws2 = wb.create_worksheet("Duplicate")
        # Implementation should handle this (either allow or modify name)
        self.assertIsNotNone(ws1)
        self.assertIsNotNone(ws2)
        
        print("All edge cases handled successfully")
    
    def test_worksheet_management_api_methods(self):
        """Test all worksheet management API methods."""
        wb = Workbook()
        
        # Test create_worksheet
        ws = wb.create_worksheet()
        self.assertIsNotNone(ws)
        
        ws = wb.create_worksheet("CustomName")
        self.assertIsNotNone(ws)
        self.assertEqual(ws.name, "CustomName")
        
        # Test remove_worksheet by index
        ws1 = wb.create_worksheet("ToRemove1")
        initial_count = len(wb.worksheets)
        wb.remove_worksheet(len(wb.worksheets) - 1)
        self.assertEqual(len(wb.worksheets), initial_count - 1)
        
        # Test remove_worksheet by name
        ws2 = wb.create_worksheet("ToRemove2")
        initial_count = len(wb.worksheets)
        wb.remove_worksheet("ToRemove2")
        self.assertEqual(len(wb.worksheets), initial_count - 1)
        
        # Test remove_worksheet by object
        ws3 = wb.create_worksheet("ToRemove3")
        initial_count = len(wb.worksheets)
        wb.remove_worksheet(ws3)
        self.assertEqual(len(wb.worksheets), initial_count - 1)
        
        # Test copy_worksheet by index
        ws_source = wb.create_worksheet("Source")
        ws_source.cells["A1"] = Cell("Test")
        ws_copy = wb.copy_worksheet(len(wb.worksheets) - 1)
        self.assertIsNotNone(ws_copy)
        self.assertEqual(ws_copy.cells["A1"].value, "Test")
        
        # Test copy_worksheet by name
        ws_copy2 = wb.copy_worksheet("Source")
        self.assertIsNotNone(ws_copy2)
        
        # Test copy_worksheet by object
        ws_copy3 = wb.copy_worksheet(ws_source)
        self.assertIsNotNone(ws_copy3)
        
        # Test get_worksheet_by_name
        found = wb.get_worksheet_by_name("Source")
        self.assertIsNotNone(found)
        self.assertEqual(found.name, "Source")
        
        # Test get_worksheet_by_index
        found = wb.get_worksheet_by_index(0)
        self.assertIsNotNone(found)
        
        # Test get_active_worksheet
        active = wb.get_active_worksheet()
        self.assertIsNotNone(active)
        
        # Test set_active_worksheet by index
        wb.set_active_worksheet(1)
        active = wb.get_active_worksheet()
        self.assertEqual(active, wb.worksheets[1])
        
        # Test set_active_worksheet by name
        wb.set_active_worksheet("Source")
        active = wb.get_active_worksheet()
        self.assertEqual(active.name, "Source")
        
        # Test set_active_worksheet by object
        ws_new = wb.create_worksheet("NewActive")
        wb.set_active_worksheet(ws_new)
        active = wb.get_active_worksheet()
        self.assertEqual(active.name, "NewActive")


if __name__ == '__main__':
    unittest.main()
