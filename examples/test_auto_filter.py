"""
Comprehensive test for AutoFilter functionality.

Tests AutoFilter persistence and loading from xlsx files.
"""

import unittest
import sys
import os
import tempfile
import zipfile
import xml.etree.ElementTree as ET

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from aspose_cells import Workbook


class TestAutoFilterPersistence(unittest.TestCase):
    """Test cases for AutoFilter persistence and loading."""
    
    def setUp(self):
        """Set up test fixtures."""
        # Create outputfiles directory if it doesn't exist
        output_dir = os.path.join(os.path.dirname(__file__), '..', 'outputfiles')
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        self.output_dir = output_dir
    
    def tearDown(self):
        """Clean up test files."""
        # Don't remove generated files - keep them for inspection
        pass
    
    def test_auto_filter_range_persistence(self):
        """Test that auto filter range is persisted correctly."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['C1'].value = "City"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 30
        ws.cells['C2'].value = "New York"
        ws.cells['A3'].value = "Bob"
        ws.cells['B3'].value = 25
        ws.cells['C3'].value = "London"
        
        # Set auto filter range
        ws.auto_filter.range = "A1:C3"
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_auto_filter_range.xlsx')
        wb.save(test_file)
        wb2 = Workbook(test_file)
        ws2 = wb2.worksheets[0]
        
        # Verify auto filter range
        self.assertEqual(ws2.auto_filter.range, "A1:C3")
    
    def test_value_filter_persistence(self):
        """Test that value filters are persisted correctly."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 30
        ws.cells['A3'].value = "Bob"
        ws.cells['B3'].value = 25
        ws.cells['A4'].value = "Charlie"
        ws.cells['B4'].value = 35
        
        # Set auto filter range and apply value filter
        ws.auto_filter.range = "A1:B4"
        ws.auto_filter.filter(0, ["Alice", "Charlie"])
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_value_filter.xlsx')
        wb.save(test_file)
        wb2 = Workbook(test_file)
        ws2 = wb2.worksheets[0]
        
        # Verify filter
        self.assertEqual(ws2.auto_filter.range, "A1:B4")
        self.assertIn(0, ws2.auto_filter.filter_columns)
        filter_col = ws2.auto_filter.filter_columns[0]
        self.assertEqual(len(filter_col.filters), 2)
        self.assertIn("Alice", filter_col.filters)
        self.assertIn("Charlie", filter_col.filters)
    
    def test_custom_filter_persistence(self):
        """Test that custom filters are persisted correctly."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 30
        ws.cells['A3'].value = "Bob"
        ws.cells['B3'].value = 25
        ws.cells['A4'].value = "Charlie"
        ws.cells['B4'].value = 35
        
        # Set auto filter range and apply custom filter
        ws.auto_filter.range = "A1:B4"
        ws.auto_filter.custom_filter(1, 'greaterThan', 25)
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_custom_filter.xlsx')
        wb.save(test_file)
        wb2 = Workbook(test_file)
        ws2 = wb2.worksheets[0]
        
        # Verify custom filter
        self.assertIn(1, ws2.auto_filter.filter_columns)
        filter_col = ws2.auto_filter.filter_columns[1]
        self.assertEqual(len(filter_col.custom_filters), 1)
        self.assertEqual(filter_col.custom_filters[0], ('greaterThan', '25'))
    
    def test_color_filter_persistence(self):
        """Test that color filters are persisted correctly."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 30
        ws.cells['A3'].value = "Bob"
        ws.cells['B3'].value = 25
        
        # Set auto filter range and apply color filter
        ws.auto_filter.range = "A1:B3"
        ws.auto_filter.filter_by_color(0, 'FFFF0000', True)
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_color_filter.xlsx')
        wb.save(test_file)
        wb2 = Workbook(test_file)
        ws2 = wb2.worksheets[0]
        
        # Verify color filter
        self.assertIn(0, ws2.auto_filter.filter_columns)
        filter_col = ws2.auto_filter.filter_columns[0]
        self.assertIsNotNone(filter_col.color_filter)
        self.assertEqual(filter_col.color_filter['color'], 'FFFF0000')
        self.assertTrue(filter_col.color_filter['cell_color'])
    
    def test_top10_filter_persistence(self):
        """Test that top10 filters are persisted correctly."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Score"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 95
        ws.cells['A3'].value = "Bob"
        ws.cells['B3'].value = 87
        ws.cells['A4'].value = "Charlie"
        ws.cells['B4'].value = 92
        
        # Set auto filter range and apply top10 filter
        ws.auto_filter.range = "A1:B4"
        ws.auto_filter.filter_top10(1, top=True, percent=False, val=3)
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_top10_filter.xlsx')
        wb.save(test_file)
        wb2 = Workbook(test_file)
        ws2 = wb2.worksheets[0]
        
        # Verify top10 filter
        self.assertIn(1, ws2.auto_filter.filter_columns)
        filter_col = ws2.auto_filter.filter_columns[1]
        self.assertIsNotNone(filter_col.top10_filter)
        self.assertTrue(filter_col.top10_filter['top'])
        self.assertFalse(filter_col.top10_filter['percent'])
        self.assertEqual(filter_col.top10_filter['val'], 3)
    
    def test_dynamic_filter_persistence(self):
        """Test that dynamic filters are persisted correctly."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Date"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = "2024-01-15"
        ws.cells['A3'].value = "Bob"
        ws.cells['B3'].value = "2024-01-10"
        
        # Set auto filter range and apply dynamic filter
        ws.auto_filter.range = "A1:B3"
        ws.auto_filter.filter_dynamic(1, 'aboveAverage')
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_dynamic_filter.xlsx')
        wb.save(test_file)
        wb2 = Workbook(test_file)
        ws2 = wb2.worksheets[0]
        
        # Verify dynamic filter
        self.assertIn(1, ws2.auto_filter.filter_columns)
        filter_col = ws2.auto_filter.filter_columns[1]
        self.assertIsNotNone(filter_col.dynamic_filter)
        self.assertEqual(filter_col.dynamic_filter['type'], 'aboveAverage')
    
    def test_sort_state_persistence(self):
        """Test that sort state is persisted correctly."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 30
        ws.cells['A3'].value = "Bob"
        ws.cells['B3'].value = 25
        
        # Set auto filter range and sort
        ws.auto_filter.range = "A1:B3"
        ws.auto_filter.sort(1, True)
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_sort_state.xlsx')
        wb.save(test_file)
        wb2 = Workbook(test_file)
        ws2 = wb2.worksheets[0]
        
        # Verify sort state (using ECMA-376 compliant structure)
        self.assertIsNotNone(ws2.auto_filter.sort_state)
        self.assertEqual(ws2.auto_filter.sort_state['column_index'], 1)
        self.assertEqual(ws2.auto_filter.sort_state['descending'], False)
    
    def test_multiple_columns_filter_persistence(self):
        """Test that filters on multiple columns are persisted correctly."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['C1'].value = "City"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 30
        ws.cells['C2'].value = "New York"
        ws.cells['A3'].value = "Bob"
        ws.cells['B3'].value = 25
        ws.cells['C3'].value = "London"
        
        # Set auto filter range and apply filters to multiple columns
        ws.auto_filter.range = "A1:C3"
        ws.auto_filter.filter(0, ["Alice"])
        ws.auto_filter.custom_filter(1, 'greaterThan', 25)
        ws.auto_filter.filter(2, ["New York"])
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_multiple_columns_filter.xlsx')
        wb.save(test_file)
        wb2 = Workbook(test_file)
        ws2 = wb2.worksheets[0]
        
        # Verify all filters
        self.assertIn(0, ws2.auto_filter.filter_columns)
        self.assertIn(1, ws2.auto_filter.filter_columns)
        self.assertIn(2, ws2.auto_filter.filter_columns)
        
        # Column 0: value filter
        filter_col_0 = ws2.auto_filter.filter_columns[0]
        self.assertEqual(len(filter_col_0.filters), 1)
        self.assertIn("Alice", filter_col_0.filters)
        
        # Column 1: custom filter
        filter_col_1 = ws2.auto_filter.filter_columns[1]
        self.assertEqual(len(filter_col_1.custom_filters), 1)
        self.assertEqual(filter_col_1.custom_filters[0], ('greaterThan', '25'))
        
        # Column 2: value filter
        filter_col_2 = ws2.auto_filter.filter_columns[2]
        self.assertEqual(len(filter_col_2.filters), 1)
        self.assertIn("New York", filter_col_2.filters)
    
    def test_filter_button_visibility_persistence(self):
        """Test that filter button visibility is persisted correctly."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 30
        
        # Set auto filter range and hide filter button for column 0
        ws.auto_filter.range = "A1:B2"
        ws.auto_filter.show_filter_button(0, False)
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_filter_button_visibility.xlsx')
        wb.save(test_file)
        wb2 = Workbook(test_file)
        ws2 = wb2.worksheets[0]
        
        # Verify filter button visibility
        self.assertIn(0, ws2.auto_filter.filter_columns)
        filter_col = ws2.auto_filter.filter_columns[0]
        self.assertFalse(filter_col.filter_button)
    
    def test_auto_filter_xml_structure(self):
        """Test that auto filter XML structure is correct."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 30
        
        # Set auto filter range and apply filter
        ws.auto_filter.range = "A1:B2"
        ws.auto_filter.filter(0, ["Alice"])
        
        # Save file
        test_file = os.path.join(self.output_dir, 'test_auto_filter_xml_structure.xlsx')
        wb.save(test_file)
        
        # Read and verify XML structure
        with zipfile.ZipFile(test_file, 'r') as zf:
            sheet_xml = zf.read('xl/worksheets/sheet1.xml')
            root = ET.fromstring(sheet_xml)
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            # Verify autoFilter element exists
            auto_filter = root.find('.//ns:autoFilter', ns)
            self.assertIsNotNone(auto_filter)
            self.assertEqual(auto_filter.attrib.get('ref'), 'A1:B2')
            
            # Verify filterColumn element exists
            filter_col = auto_filter.find('ns:filterColumn', ns)
            self.assertIsNotNone(filter_col)
            self.assertEqual(int(filter_col.attrib.get('colId')), 0)
            
            # Verify filters element exists
            filters = filter_col.find('ns:filters', ns)
            self.assertIsNotNone(filters)
            
            # Verify filter element exists
            filter_elem = filters.find('ns:filter', ns)
            self.assertIsNotNone(filter_elem)
            self.assertEqual(filter_elem.attrib.get('val'), 'Alice')
    
    def test_clear_filters_persistence(self):
        """Test that clearing filters is persisted correctly."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 30
        
        # Set auto filter range and apply filter
        ws.auto_filter.range = "A1:B2"
        ws.auto_filter.filter(0, ["Alice"])
        
        # Clear filters
        ws.auto_filter.clear_all_filters()
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_clear_filters.xlsx')
        wb.save(test_file)
        wb2 = Workbook(test_file)
        ws2 = wb2.worksheets[0]
        
        # Verify filters are cleared
        self.assertEqual(ws2.auto_filter.range, "A1:B2")
        self.assertEqual(ws2.auto_filter.filter_columns, {})
    
    def test_remove_auto_filter_persistence(self):
        """Test that removing auto filter is persisted correctly."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 30
        
        # Set auto filter range and apply filter
        ws.auto_filter.range = "A1:B2"
        ws.auto_filter.filter(0, ["Alice"])
        
        # Remove auto filter
        ws.auto_filter.remove()
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_remove_auto_filter.xlsx')
        wb.save(test_file)
        wb2 = Workbook(test_file)
        ws2 = wb2.worksheets[0]
        
        # Verify auto filter is removed
        self.assertIsNone(ws2.auto_filter.range)
        self.assertEqual(ws2.auto_filter.filter_columns, {})
    
    def test_numeric_value_filter_persistence(self):
        """Test that numeric value filters are persisted correctly."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "ID"
        ws.cells['B1'].value = "Score"
        ws.cells['A2'].value = 1
        ws.cells['B2'].value = 95
        ws.cells['A3'].value = 2
        ws.cells['B3'].value = 87
        
        # Set auto filter range and apply numeric filter
        ws.auto_filter.range = "A1:B3"
        ws.auto_filter.filter(0, [1, 2])
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_numeric_value_filter.xlsx')
        wb.save(test_file)
        wb2 = Workbook(test_file)
        ws2 = wb2.worksheets[0]
        
        # Verify numeric filter
        self.assertIn(0, ws2.auto_filter.filter_columns)
        filter_col = ws2.auto_filter.filter_columns[0]
        self.assertEqual(len(filter_col.filters), 2)
        self.assertIn("1", filter_col.filters)
        self.assertIn("2", filter_col.filters)
    
    def test_set_range_with_indices_persistence(self):
        """Test that setting range with indices is persisted correctly."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add some data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 30
        
        # Set auto filter range using indices
        ws.auto_filter.set_range(1, 1, 2, 2)
        
        # Save and reload
        test_file = os.path.join(self.output_dir, 'test_set_range_with_indices.xlsx')
        wb.save(test_file)
        wb2 = Workbook(test_file)
        ws2 = wb2.worksheets[0]
        
        # Verify range
        self.assertEqual(ws2.auto_filter.range, "A1:B2")


if __name__ == '__main__':
    unittest.main()
