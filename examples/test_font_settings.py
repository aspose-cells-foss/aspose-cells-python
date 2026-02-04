import unittest
import os
import sys

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose_cells import Workbook, Worksheet, Cell, Style, Font


class TestFontSettings(unittest.TestCase):
    """
    Test comprehensive font settings functionality.
    
    Features tested:
    - Font name setting (Arial, Times New Roman, Courier New)
    - Font size setting (8-72 points)
    - Font color setting with hex values (black, white, red, green, blue, yellow, magenta, cyan, orange, purple)
    - Font bold setting
    - Font italic setting
    - Font underline setting
    - Font strikethrough setting
    - Combined font styles (multiple attributes together)
    - Font default values (Calibri, 11pt, black, no styles)
    - Font copying
    - Save font settings to file
    """
    
    def setUp(self):
        """Set up test workbook and worksheet."""
        self.workbook = Workbook()
        self.worksheet = self.workbook.worksheets[0]
    
    def test_font_name(self):
        """Test font name setting."""
        cell = Cell("Test")
        cell.style.font.name = "Arial"
        self.worksheet.cells["A1"] = cell
        
        cell2 = Cell("Test")
        cell2.style.font.name = "Times New Roman"
        self.worksheet.cells["A2"] = cell2
        
        cell3 = Cell("Test")
        cell3.style.font.name = "Courier New"
        self.worksheet.cells["A3"] = cell3
    
    def test_font_size(self):
        """Test font size setting."""
        sizes = [8, 10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72]
        for i, size in enumerate(sizes):
            cell = Cell(f"Size {size}")
            cell.style.font.size = size
            self.worksheet.cells[f"A{i+1}"] = cell
    
    def test_font_color_hex(self):
        """Test font color with hex values."""
        colors = [
            "FF000000",  # Black
            "FFFFFFFF",  # White
            "FFFF0000",  # Red
            "FF00FF00",  # Green
            "FF0000FF",  # Blue
            "FFFFFF00",  # Yellow
            "FFFF00FF",  # Magenta
            "FF00FFFF",  # Cyan
            "FFFFA500",  # Orange
            "FF800080",  # Purple
        ]
        
        for i, color in enumerate(colors):
            cell = Cell(f"Color {color}")
            cell.style.font.color = color
            self.worksheet.cells[f"A{i+1}"] = cell
    
    def test_font_color_rgb(self):
        """Test font color with RGB hex values."""
        colors = [
            ("FF000000", "Black"),
            ("FFFF0000", "Red"),
            ("FF0000FF", "Blue"),
            ("FF00FF00", "Green"),
            ("FFFFFF00", "Yellow"),
            ("FFFFFFFF", "White")
        ]
        
        for i, (color, name) in enumerate(colors):
            cell = Cell(f"Color {name}")
            cell.style.font.color = color
            self.worksheet.cells[f"A{i+1}"] = cell
    
    def test_font_bold(self):
        """Test font bold setting."""
        cell = Cell("Bold Text")
        cell.style.font.bold = True
        self.worksheet.cells["A1"] = cell
        
        cell2 = Cell("Normal Text")
        cell2.style.font.bold = False
        self.worksheet.cells["A2"] = cell2
    
    def test_font_italic(self):
        """Test font italic setting."""
        cell = Cell("Italic Text")
        cell.style.font.italic = True
        self.worksheet.cells["A1"] = cell
        
        cell2 = Cell("Normal Text")
        cell2.style.font.italic = False
        self.worksheet.cells["A2"] = cell2
    
    def test_font_underline(self):
        """Test font underline setting."""
        cell = Cell("Underlined Text")
        cell.style.font.underline = True
        self.worksheet.cells["A1"] = cell
        
        cell2 = Cell("Normal Text")
        cell2.style.font.underline = False
        self.worksheet.cells["A2"] = cell2
    
    def test_font_strikethrough(self):
        """Test font strikethrough setting."""
        cell = Cell("Strikethrough Text")
        cell.style.font.strikethrough = True
        self.worksheet.cells["A1"] = cell
        
        cell2 = Cell("Normal Text")
        cell2.style.font.strikethrough = False
        self.worksheet.cells["A2"] = cell2
    
    def test_font_combined_styles(self):
        """Test font with multiple style attributes."""
        cell = Cell("Combined Styles")
        cell.style.font.name = "Arial"
        cell.style.font.size = 14
        cell.style.font.color = "FF0000FF"  # Blue
        cell.style.font.bold = True
        cell.style.font.italic = True
        cell.style.font.underline = True
        cell.style.font.strikethrough = True
        self.worksheet.cells["A1"] = cell
    
    def test_font_default_values(self):
        """Test font default values."""
        font = Font()
        self.assertEqual(font.name, 'Calibri')
        self.assertEqual(font.size, 11)
        self.assertEqual(font.color, 'FF000000')
        self.assertFalse(font.bold)
        self.assertFalse(font.italic)
        self.assertFalse(font.underline)
        self.assertFalse(font.strikethrough)
    
    def test_font_copy(self):
        """Test font copying."""
        original_font = Font(
            name="Arial",
            size=12,
            color="FF0000FF",
            bold=True,
            italic=True,
            underline=True,
            strikethrough=True
        )
        
        new_font = Font(**vars(original_font))
        
        self.assertEqual(new_font.name, original_font.name)
        self.assertEqual(new_font.size, original_font.size)
        self.assertEqual(new_font.color, original_font.color)
        self.assertEqual(new_font.bold, original_font.bold)
        self.assertEqual(new_font.italic, original_font.italic)
        self.assertEqual(new_font.underline, original_font.underline)
        self.assertEqual(new_font.strikethrough, original_font.strikethrough)
    
    def test_comprehensive_font_test(self):
        """Test all font settings comprehensively."""
        # Test data for comprehensive font testing
        test_cases = [
            {
                'text': 'Arial 10 Black Normal',
                'font': Font(name='Arial', size=10, color='FF000000', bold=False, italic=False, underline=False, strikethrough=False)
            },
            {
                'text': 'Arial 12 Red Bold',
                'font': Font(name='Arial', size=12, color='FF0000FF', bold=True, italic=False, underline=False, strikethrough=False)
            },
            {
                'text': 'Times 14 Blue Italic',
                'font': Font(name='Times New Roman', size=14, color='FF0000FF', bold=False, italic=True, underline=False, strikethrough=False)
            },
            {
                'text': 'Courier 16 Green Bold Italic',
                'font': Font(name='Courier New', size=16, color='FF00FF00', bold=True, italic=True, underline=False, strikethrough=False)
            },
            {
                'text': 'Arial 18 Yellow Underline',
                'font': Font(name='Arial', size=18, color='FFFFFF00', bold=False, italic=False, underline=True, strikethrough=False)
            },
            {
                'text': 'Arial 20 Orange Strikethrough',
                'font': Font(name='Arial', size=20, color='FFFFA500', bold=False, italic=False, underline=False, strikethrough=True)
            },
            {
                'text': 'Arial 22 Purple Bold Italic Underline',
                'font': Font(name='Arial', size=22, color='FF800080', bold=True, italic=True, underline=True, strikethrough=False)
            },
            {
                'text': 'Arial 24 Cyan Bold Italic Strikethrough',
                'font': Font(name='Arial', size=24, color='FF00FFFF', bold=True, italic=True, underline=False, strikethrough=True)
            },
            {
                'text': 'Arial 26 Magenta All Styles',
                'font': Font(name='Arial', size=26, color='FFFF00FF', bold=True, italic=True, underline=True, strikethrough=True)
            },
            {
                'text': 'Arial 28 White Bold',
                'font': Font(name='Arial', size=28, color='FFFFFFFF', bold=True, italic=False, underline=False, strikethrough=False)
            }
        ]
        
        # Apply all test cases to cells
        for i, test_case in enumerate(test_cases):
            cell = Cell(test_case['text'])
            cell.style.font = test_case['font']
            self.worksheet.cells[f"A{i+1}"] = cell
    
    def test_save_font_settings(self):
        """Test saving font settings to file."""
        # Create comprehensive test data
        self.test_comprehensive_font_test()
        
        # Save to outputfiles folder
        output_path = 'outputfiles/test_font_settings.xlsx'
        
        # Ensure outputfiles directory exists
        os.makedirs('outputfiles', exist_ok=True)
        
        self.workbook.save(output_path)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify file is not empty
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"Font settings test file saved to: {output_path}")
        print(f"File size: {file_size} bytes")


if __name__ == '__main__':
    unittest.main()