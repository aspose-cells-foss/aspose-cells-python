"""
Test Suite for XLSX to JSON Conversion

This test suite covers converting Excel files to JSON format including:
- Basic XLSX to JSON conversion
- Loading comprehensive sales report data
- Exporting to JSON format
"""

import unittest
import os
import sys
import json

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose_cells import Workbook, SaveFormat
from aspose_cells.json_handler import JsonHandler, JsonSaveOptions


class TestXLSXToJSONConversion(unittest.TestCase):
    """Test XLSX to JSON conversion functionality."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_sales_report_to_json(self):
        """Test converting comprehensive sales report to JSON."""
        # Load Excel file
        input_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 
                                'input/sales_report_comprehensive.xlsx')
        self.assertTrue(os.path.exists(input_path), 
                       f"Input file {input_path} does not exist")
        
        wb = Workbook(input_path)
        
        # Verify workbook loaded
        self.assertGreater(len(wb.worksheets), 0, 
                          "Workbook should have at least one worksheet")
        
        # Export to JSON
        output_path = os.path.join(self.test_dir, 'sales_report_comprehensive.json')
        wb.save_as_json(output_path)
        
        # Verify JSON file was created
        self.assertTrue(os.path.exists(output_path), 
                       f"JSON file {output_path} was not created")
        
        # Verify JSON file is valid
        with open(output_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        
        # Verify JSON structure
        self.assertIn('worksheets', json_data, 
                     "JSON should contain 'worksheets' key")
        self.assertIsInstance(json_data['worksheets'], list, 
                          "'worksheets' should be a list")
        
        # Verify each worksheet has expected structure
        for sheet in json_data['worksheets']:
            self.assertIn('name', sheet, 
                         "Each worksheet should have 'name'")
            self.assertIn('data', sheet, 
                         "Each worksheet should have 'data'")
            self.assertIsInstance(sheet['data'], list, 
                              "'data' should be a list")
        
        # Print summary
        print(f"\nSuccessfully converted {input_path} to {output_path}")
        print(f"Number of worksheets: {len(wb.worksheets)}")
        for i, ws in enumerate(wb.worksheets):
            print(f"  Worksheet {i}: {ws.name}")
        print(f"JSON file size: {os.path.getsize(output_path)} bytes")
    
    def test_sales_report_to_json_with_options(self):
        """Test converting sales report to JSON with custom options."""
        # Load Excel file
        input_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 
                                'input/sales_report_comprehensive.xlsx')
        wb = Workbook(input_path)
        
        # Create custom options
        options = JsonSaveOptions()
        options.include_worksheet_name = True
        options.indent = 4
        options.skip_empty_rows = True
        options.empty_cell_value = ""
        
        # Export to JSON with options
        output_path = os.path.join(self.test_dir, 'sales_report_comprehensive_custom.json')
        wb.save_as_json(output_path, options)
        
        # Verify JSON file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify JSON file is valid and has proper indentation
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
            json_data = json.loads(content)
        
        # Verify indentation (4 spaces)
        self.assertIn('    ', content, 
                     "JSON should be indented with 4 spaces")
        
        # Verify structure
        self.assertIn('worksheets', json_data)
        
        print(f"\nSuccessfully converted with custom options to {output_path}")
        print(f"JSON file size: {os.path.getsize(output_path)} bytes")
    
    def test_sales_report_to_json_single_worksheet(self):
        """Test converting only first worksheet to JSON."""
        # Load Excel file
        input_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 
                                'input/sales_report_comprehensive.xlsx')
        wb = Workbook(input_path)
        
        # Create options to export only first worksheet
        options = JsonSaveOptions()
        options.worksheet_index = 0
        
        # Export to JSON
        output_path = os.path.join(self.test_dir, 'sales_report_sheet0.json')
        wb.save_as_json(output_path, options)
        
        # Verify JSON file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify only one worksheet is exported
        with open(output_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        
        self.assertEqual(len(json_data['worksheets']), 1, 
                        "Only one worksheet should be exported")
        self.assertEqual(json_data['worksheets'][0]['name'], 
                        wb.worksheets[0].name,
                        "First worksheet name should match")
        
        print(f"\nSuccessfully exported first worksheet to {output_path}")
    
    def test_sales_report_to_json_using_save_format(self):
        """Test converting using SaveFormat.JSON."""
        # Load Excel file
        input_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 
                                'input/sales_report_comprehensive.xlsx')
        wb = Workbook(input_path)
        
        # Export to JSON using SaveFormat
        output_path = os.path.join(self.test_dir, 'sales_report_using_save_format.json')
        wb.save(output_path, SaveFormat.JSON)
        
        # Verify JSON file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify JSON file is valid
        with open(output_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        
        self.assertIn('worksheets', json_data)
        
        print(f"\nSuccessfully converted using SaveFormat.JSON to {output_path}")


if __name__ == '__main__':
    unittest.main()
