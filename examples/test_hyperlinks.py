"""
Test Hyperlinks Feature

This test verifies that hyperlinks are correctly created, saved to,
and loaded from XLSX files according to ECMA-376 specification.
"""

import unittest
import os
import sys
import zipfile
import xml.etree.ElementTree as ET

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from aspose_cells import Workbook


class TestHyperlinks(unittest.TestCase):
    """Test cases for hyperlinks feature."""

    def setUp(self):
        """Set up test fixtures."""
        self.ns = {
            'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }
        self.rels_ns = {
            'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'
        }

    def test_add_external_hyperlink(self):
        """Test adding an external web hyperlink."""
        print("\n" + "="*70)
        print("Test: Add External Hyperlink")
        print("="*70)

        # Create workbook
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "Hyperlinks"

        # Add external hyperlink
        print("\nAdding external hyperlink to A1...")
        link = ws.hyperlinks.add("A1", "https://www.example.com")
        link.text_to_display = "Visit Example"
        link.screen_tip = "Click to visit example.com"

        # Add cell value for display
        ws.cells['A1'].value = "Visit Example"

        # Verify hyperlink properties
        self.assertEqual(link.range, "A1")
        self.assertEqual(link.address, "https://www.example.com")
        self.assertEqual(link.text_to_display, "Visit Example")
        self.assertEqual(link.screen_tip, "Click to visit example.com")
        self.assertEqual(link.type, "External")
        print("  [OK] Hyperlink created successfully")

        # Verify collection
        self.assertEqual(ws.hyperlinks.count, 1)
        print(f"  [OK] Collection has {ws.hyperlinks.count} hyperlink(s)")

        # Save to file
        output_file = "outputfiles/test_hyperlink_external.xlsx"
        print(f"\nSaving to {output_file}...")
        wb.save(output_file)
        self.assertTrue(os.path.exists(output_file))
        print(f"  [OK] File saved successfully")

        print("="*70 + "\n")

    def test_add_internal_hyperlink(self):
        """Test adding an internal hyperlink to another sheet."""
        print("\n" + "="*70)
        print("Test: Add Internal Hyperlink")
        print("="*70)

        # Create workbook with two sheets
        wb = Workbook()
        ws1 = wb.worksheets[0]
        ws1.name = "Sheet1"
        from aspose_cells import Worksheet
        ws2 = Worksheet("Sheet2")
        wb.worksheets.append(ws2)

        # Add internal hyperlink
        print("\nAdding internal hyperlink to A1...")
        link = ws1.hyperlinks.add("A1", sub_address="Sheet2!A1")
        link.text_to_display = "Go to Sheet2"
        link.screen_tip = "Navigate to Sheet2"

        # Add cell values for display
        ws1.cells['A1'].value = "Go to Sheet2"
        ws2.cells['A1'].value = "Welcome to Sheet2!"

        # Verify hyperlink properties
        self.assertEqual(link.range, "A1")
        self.assertEqual(link.sub_address, "Sheet2!A1")
        self.assertEqual(link.text_to_display, "Go to Sheet2")
        self.assertEqual(link.type, "Internal")
        print("  [OK] Internal hyperlink created")

        # Save to file
        output_file = "outputfiles/test_hyperlink_internal.xlsx"
        print(f"\nSaving to {output_file}...")
        wb.save(output_file)
        self.assertTrue(os.path.exists(output_file))
        print(f"  [OK] File saved successfully")

        print("="*70 + "\n")

    def test_hyperlink_roundtrip(self):
        """Test that hyperlinks survive save/load roundtrip."""
        print("\n" + "="*70)
        print("Test: Hyperlink Roundtrip (Save/Load)")
        print("="*70)

        # Create workbook with hyperlinks
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "Links"

        # Add different types of hyperlinks
        print("\nAdding various hyperlink types...")
        link1 = ws.hyperlinks.add("A1", "https://www.example.com")
        link1.text_to_display = "Example Website"
        link1.screen_tip = "Visit example.com"
        print("  Added: External HTTPS link at A1")

        link2 = ws.hyperlinks.add("A2", "mailto:user@example.com")
        link2.text_to_display = "Email Us"
        link2.screen_tip = "Send us an email"
        print("  Added: Email link at A2")

        link3 = ws.hyperlinks.add("A3", "ftp://ftp.example.com/files")
        link3.text_to_display = "FTP Server"
        print("  Added: FTP link at A3")

        link4 = ws.hyperlinks.add("A4", sub_address="Links!A1")
        link4.text_to_display = "Go to Top"
        link4.screen_tip = "Jump to cell A1"
        print("  Added: Internal link at A4")

        # Add cell values
        ws.cells['A1'].value = "Example Website"
        ws.cells['A2'].value = "Email Us"
        ws.cells['A3'].value = "FTP Server"
        ws.cells['A4'].value = "Go to Top"

        # Save
        output_file = "outputfiles/test_hyperlinks_roundtrip.xlsx"
        print(f"\nSaving to {output_file}...")
        wb.save(output_file)
        self.assertTrue(os.path.exists(output_file))

        # Verify XML structure
        print("\nVerifying XML structure...")
        with zipfile.ZipFile(output_file, 'r') as zf:
            # Check worksheet XML
            sheet_xml = zf.read('xl/worksheets/sheet1.xml').decode('utf-8')
            print(f"  Worksheet XML length: {len(sheet_xml)} bytes")

            # Parse and verify hyperlinks element
            root = ET.fromstring(sheet_xml)
            hyperlinks_elem = root.find('main:hyperlinks', self.ns)
            self.assertIsNotNone(hyperlinks_elem, "hyperlinks element should exist")

            hyperlink_elems = hyperlinks_elem.findall('main:hyperlink', self.ns)
            self.assertEqual(len(hyperlink_elems), 4, "Should have 4 hyperlinks")
            print(f"  [OK] Found {len(hyperlink_elems)} hyperlink elements")

            # Verify first hyperlink (external HTTPS)
            link1_elem = hyperlink_elems[0]
            self.assertEqual(link1_elem.get('ref'), 'A1')
            self.assertIsNotNone(link1_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'))
            self.assertEqual(link1_elem.get('display'), 'Example Website')
            self.assertEqual(link1_elem.get('tooltip'), 'Visit example.com')
            print("  [OK] First hyperlink (HTTPS) attributes correct")

            # Verify internal hyperlink (A4)
            link4_elem = hyperlink_elems[3]
            self.assertEqual(link4_elem.get('ref'), 'A4')
            self.assertEqual(link4_elem.get('location'), 'Links!A1')
            self.assertEqual(link4_elem.get('display'), 'Go to Top')
            self.assertEqual(link4_elem.get('tooltip'), 'Jump to cell A1')
            print("  [OK] Internal hyperlink attributes correct")

            # Check relationships file
            rels_xml = zf.read('xl/worksheets/_rels/sheet1.xml.rels').decode('utf-8')
            rels_root = ET.fromstring(rels_xml)
            rel_elems = rels_root.findall('rel:Relationship', self.rels_ns)

            hyperlink_rels = [rel for rel in rel_elems
                            if 'hyperlink' in rel.get('Type', '')]
            self.assertEqual(len(hyperlink_rels), 3, "Should have 3 external hyperlink relationships")
            print(f"  [OK] Found {len(hyperlink_rels)} external hyperlink relationships")

            # Verify relationship targets
            targets = [rel.get('Target') for rel in hyperlink_rels]
            self.assertIn('https://www.example.com', targets)
            self.assertIn('mailto:user@example.com', targets)
            self.assertIn('ftp://ftp.example.com/files', targets)
            print("  [OK] Relationship targets correct (https, mailto, ftp)")

        # Load back
        print("\nLoading workbook back...")
        wb_loaded = Workbook(output_file)
        ws_loaded = wb_loaded.worksheets[0]

        # Verify hyperlinks loaded correctly
        print("Verifying loaded hyperlinks...")
        self.assertEqual(ws_loaded.hyperlinks.count, 4, "Should load 4 hyperlinks")

        # Find hyperlinks by range
        loaded_links = list(ws_loaded.hyperlinks)
        link1_loaded = next((l for l in loaded_links if l.range == 'A1'), None)
        link2_loaded = next((l for l in loaded_links if l.range == 'A2'), None)
        link3_loaded = next((l for l in loaded_links if l.range == 'A3'), None)
        link4_loaded = next((l for l in loaded_links if l.range == 'A4'), None)

        self.assertIsNotNone(link1_loaded)
        self.assertIsNotNone(link2_loaded)
        self.assertIsNotNone(link3_loaded)
        self.assertIsNotNone(link4_loaded)

        # Verify link1 (HTTPS)
        self.assertEqual(link1_loaded.address, 'https://www.example.com')
        self.assertEqual(link1_loaded.text_to_display, 'Example Website')
        self.assertEqual(link1_loaded.screen_tip, 'Visit example.com')
        self.assertEqual(link1_loaded.type, 'External')
        print("  [OK] Link 1 (HTTPS) loaded correctly")

        # Verify link2 (Email)
        self.assertEqual(link2_loaded.address, 'mailto:user@example.com')
        self.assertEqual(link2_loaded.text_to_display, 'Email Us')
        self.assertEqual(link2_loaded.screen_tip, 'Send us an email')
        self.assertEqual(link2_loaded.type, 'External')
        print("  [OK] Link 2 (Email) loaded correctly")

        # Verify link3 (FTP)
        self.assertEqual(link3_loaded.address, 'ftp://ftp.example.com/files')
        self.assertEqual(link3_loaded.text_to_display, 'FTP Server')
        self.assertEqual(link3_loaded.type, 'External')
        print("  [OK] Link 3 (FTP) loaded correctly")

        # Verify link4 (Internal)
        self.assertEqual(link4_loaded.sub_address, 'Links!A1')
        self.assertEqual(link4_loaded.text_to_display, 'Go to Top')
        self.assertEqual(link4_loaded.screen_tip, 'Jump to cell A1')
        self.assertEqual(link4_loaded.type, 'Internal')
        print("  [OK] Link 4 (Internal) loaded correctly")

        print("\n[OK] All 4 hyperlink types preserved through roundtrip!")
        print("="*70 + "\n")

    def test_multiple_hyperlinks(self):
        """Test adding multiple hyperlinks to a worksheet."""
        print("\n" + "="*70)
        print("Test: Multiple Hyperlinks")
        print("="*70)

        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "Links"

        # Add multiple hyperlinks
        print("\nAdding 5 hyperlinks...")
        ws.hyperlinks.add("A1", "https://www.google.com", text_to_display="Google")
        ws.hyperlinks.add("A2", "https://www.github.com", text_to_display="GitHub")
        ws.hyperlinks.add("A3", "https://www.python.org", text_to_display="Python")
        ws.hyperlinks.add("B1", "mailto:info@example.com", text_to_display="Contact")
        ws.hyperlinks.add("B2", sub_address="Sheet1!A1", text_to_display="Top")

        # Add cell values for display
        ws.cells['A1'].value = "Google"
        ws.cells['A2'].value = "GitHub"
        ws.cells['A3'].value = "Python"
        ws.cells['B1'].value = "Contact"
        ws.cells['B2'].value = "Top"

        self.assertEqual(ws.hyperlinks.count, 5)
        print(f"  [OK] Added {ws.hyperlinks.count} hyperlinks")

        # Iterate over hyperlinks
        print("\nHyperlinks:")
        for i, link in enumerate(ws.hyperlinks):
            print(f"  {i+1}. {link.range}: {link.text_to_display or '(no display)'}")

        # Save to file
        output_file = "outputfiles/test_hyperlinks_multiple.xlsx"
        print(f"\nSaving to {output_file}...")
        wb.save(output_file)
        self.assertTrue(os.path.exists(output_file))
        print(f"  [OK] File saved successfully")

        print("="*70 + "\n")

    def test_delete_hyperlink(self):
        """Test deleting hyperlinks."""
        print("\n" + "="*70)
        print("Test: Delete Hyperlink")
        print("="*70)

        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "DeleteTest"

        # Add hyperlinks
        ws.hyperlinks.add("A1", "https://www.example1.com", text_to_display="Link 1")
        ws.hyperlinks.add("A2", "https://www.example2.com", text_to_display="Link 2")
        ws.hyperlinks.add("A3", "https://www.example3.com", text_to_display="Link 3")
        ws.hyperlinks.add("A4", "https://www.example4.com", text_to_display="Link 4")

        # Add cell values
        ws.cells['A1'].value = "Link 1"
        ws.cells['A2'].value = "Link 2"
        ws.cells['A3'].value = "Link 3"
        ws.cells['A4'].value = "Link 4 (will remain)"

        self.assertEqual(ws.hyperlinks.count, 4)
        print(f"Initial count: {ws.hyperlinks.count}")

        # Delete by index (delete second link - A2)
        ws.hyperlinks.delete(index=1)
        self.assertEqual(ws.hyperlinks.count, 3)
        print(f"After delete by index: {ws.hyperlinks.count}")

        # Delete by object (delete first remaining link - A1)
        link = ws.hyperlinks[0]
        ws.hyperlinks.delete(hyperlink=link)
        self.assertEqual(ws.hyperlinks.count, 2)
        print(f"After delete by object: {ws.hyperlinks.count}")

        # Delete one more (delete A3)
        ws.hyperlinks.delete(index=0)
        self.assertEqual(ws.hyperlinks.count, 1)
        print(f"After another delete: {ws.hyperlinks.count}")

        # Save file with remaining hyperlink (A4)
        output_file = "outputfiles/test_hyperlinks_delete.xlsx"
        print(f"\nSaving to {output_file} (1 hyperlink remaining)...")
        wb.save(output_file)
        self.assertTrue(os.path.exists(output_file))
        print(f"  [OK] File saved successfully")
        print(f"  Remaining hyperlink: {ws.hyperlinks[0].range} - {ws.hyperlinks[0].text_to_display}")

        # Test clear all on a new workbook
        wb2 = Workbook()
        ws2 = wb2.worksheets[0]
        ws2.hyperlinks.add("B1", "https://www.test.com")
        ws2.hyperlinks.clear()
        self.assertEqual(ws2.hyperlinks.count, 0)
        print(f"\nAfter clear on new workbook: {ws2.hyperlinks.count}")

        print("="*70 + "\n")

    def test_hyperlink_validation(self):
        """Test hyperlink validation."""
        print("\n" + "="*70)
        print("Test: Hyperlink Validation")
        print("="*70)

        wb = Workbook()
        ws = wb.worksheets[0]

        # Test: Cannot specify both address and sub_address
        print("\nTest: Cannot specify both address and sub_address...")
        with self.assertRaises(ValueError) as context:
            ws.hyperlinks.add("A1", "https://www.example.com", "Sheet2!A1")
        print(f"  [OK] Raised ValueError: {context.exception}")

        # Test: Must specify either address or sub_address
        print("\nTest: Must specify either address or sub_address...")
        with self.assertRaises(ValueError) as context:
            ws.hyperlinks.add("A1")
        print(f"  [OK] Raised ValueError: {context.exception}")

        print("="*70 + "\n")

    def test_hyperlink_element_order(self):
        """Test that hyperlinks appear in correct order in XML."""
        print("\n" + "="*70)
        print("Test: Hyperlink Element Order (ECMA-376)")
        print("="*70)

        wb = Workbook()
        ws = wb.worksheets[0]

        # Add some data and features
        ws.cells['A1'].value = "Data"
        ws.hyperlinks.add("B1", "https://www.example.com")

        # Save
        output_file = "outputfiles/test_hyperlink_order.xlsx"
        print(f"\nSaving to {output_file}...")
        wb.save(output_file)

        # Verify element order
        print("Verifying element order...")
        with zipfile.ZipFile(output_file, 'r') as zf:
            sheet_xml = zf.read('xl/worksheets/sheet1.xml').decode('utf-8')
            root = ET.fromstring(sheet_xml)

            # Get all child elements
            children = list(root)
            element_names = [child.tag.split('}')[-1] for child in children]

            print(f"\nElement order: {element_names}")

            # Verify hyperlinks comes after sheetData
            if 'sheetData' in element_names and 'hyperlinks' in element_names:
                data_index = element_names.index('sheetData')
                hyperlinks_index = element_names.index('hyperlinks')
                self.assertLess(data_index, hyperlinks_index,
                              "sheetData must come before hyperlinks")
                print(f"  sheetData at index {data_index}")
                print(f"  hyperlinks at index {hyperlinks_index}")
                print("  [OK] Order is correct")

        print("\n[OK] Element order complies with ECMA-376")
        print("="*70 + "\n")

    def test_file_and_unc_hyperlinks(self):
        """Test file path and UNC path hyperlinks."""
        print("\n" + "="*70)
        print("Test: File and UNC Path Hyperlinks")
        print("="*70)

        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "FilePaths"

        # Add file path hyperlinks (Windows style)
        print("\nAdding file path hyperlinks...")
        link1 = ws.hyperlinks.add("A1", "file:///C:/Documents/report.pdf")
        link1.text_to_display = "Monthly Report"
        link1.screen_tip = "Open report in PDF viewer"
        print("  Added: Local file path (C:) at A1")

        link2 = ws.hyperlinks.add("A2", "file:///C:/Users/Documents/data.xlsx")
        link2.text_to_display = "Data File"
        print("  Added: Local Excel file at A2")

        # Add UNC path hyperlink
        link3 = ws.hyperlinks.add("A3", "file://server/share/document.docx")
        link3.text_to_display = "Network Document"
        link3.screen_tip = "Access network share"
        print("  Added: UNC network path at A3")

        # Add relative file path
        link4 = ws.hyperlinks.add("A4", "file:///./resources/config.json")
        link4.text_to_display = "Config File"
        print("  Added: Relative file path at A4")

        # Add cell values
        ws.cells['A1'].value = "Monthly Report"
        ws.cells['A2'].value = "Data File"
        ws.cells['A3'].value = "Network Document"
        ws.cells['A4'].value = "Config File"

        self.assertEqual(ws.hyperlinks.count, 4)
        print(f"\n[OK] Added {ws.hyperlinks.count} file/UNC hyperlinks")

        # Save to file
        output_file = "outputfiles/test_hyperlinks_files.xlsx"
        print(f"\nSaving to {output_file}...")
        wb.save(output_file)
        self.assertTrue(os.path.exists(output_file))
        print(f"  [OK] File saved successfully")

        # Load and verify
        print("\nLoading back to verify...")
        wb_loaded = Workbook(output_file)
        ws_loaded = wb_loaded.worksheets[0]
        self.assertEqual(ws_loaded.hyperlinks.count, 4)

        # Verify each link type
        loaded_links = list(ws_loaded.hyperlinks)
        for link in loaded_links:
            self.assertTrue(link.address.startswith("file://"))
            print(f"  [OK] {link.range}: {link.text_to_display}")

        print("\n[OK] All file path hyperlinks working correctly!")
        print("="*70 + "\n")

    def test_comprehensive_mixed_hyperlinks(self):
        """Test a worksheet with all hyperlink types mixed together."""
        print("\n" + "="*70)
        print("Test: Comprehensive Mixed Hyperlink Types")
        print("="*70)

        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "ComprehensiveTest"

        # Add a variety of hyperlinks in a realistic scenario
        print("\nCreating comprehensive hyperlink test sheet...")

        # Header
        ws.cells['A1'].value = "Hyperlink Type"
        ws.cells['B1'].value = "Link"
        ws.cells['C1'].value = "Description"

        # Web links (HTTPS, HTTP)
        ws.cells['A2'].value = "Web (HTTPS)"
        ws.hyperlinks.add("B2", "https://www.example.com", text_to_display="Example Site")
        ws.cells['B2'].value = "Example Site"
        ws.cells['C2'].value = "Secure website link"

        ws.cells['A3'].value = "Web (HTTP)"
        ws.hyperlinks.add("B3", "http://legacy.example.com", text_to_display="Legacy Site")
        ws.cells['B3'].value = "Legacy Site"
        ws.cells['C3'].value = "Non-secure legacy site"

        # Email links
        ws.cells['A4'].value = "Email (Simple)"
        ws.hyperlinks.add("B4", "mailto:contact@example.com", text_to_display="Contact Email")
        ws.cells['B4'].value = "Contact Email"
        ws.cells['C4'].value = "Send email to contact"

        ws.cells['A5'].value = "Email (Subject)"
        ws.hyperlinks.add("B5", "mailto:support@example.com?subject=Help%20Request",
                          text_to_display="Support Email")
        ws.cells['B5'].value = "Support Email"
        ws.cells['C5'].value = "Email with preset subject"

        # File links
        ws.cells['A6'].value = "Local File"
        ws.hyperlinks.add("B6", "file:///C:/Reports/annual.pdf", text_to_display="Annual Report")
        ws.cells['B6'].value = "Annual Report"
        ws.cells['C6'].value = "Open local PDF file"

        ws.cells['A7'].value = "Network File"
        ws.hyperlinks.add("B7", "file://server/shared/data.xlsx", text_to_display="Shared Data")
        ws.cells['B7'].value = "Shared Data"
        ws.cells['C7'].value = "Access network share"

        # FTP link
        ws.cells['A8'].value = "FTP Server"
        ws.hyperlinks.add("B8", "ftp://ftp.example.com/files/", text_to_display="FTP Files")
        ws.cells['B8'].value = "FTP Files"
        ws.cells['C8'].value = "Browse FTP directory"

        # Internal links
        ws.cells['A9'].value = "Internal (Same Sheet)"
        ws.hyperlinks.add("B9", sub_address="ComprehensiveTest!A1", text_to_display="Go to Top")
        ws.cells['B9'].value = "Go to Top"
        ws.cells['C9'].value = "Jump to cell A1"

        ws.cells['A10'].value = "Internal (Named Range)"
        ws.hyperlinks.add("B10", sub_address="ComprehensiveTest!B2:B9",
                          text_to_display="View All Links")
        ws.cells['B10'].value = "View All Links"
        ws.cells['C10'].value = "Select range of links"

        total_links = ws.hyperlinks.count
        self.assertEqual(total_links, 9)
        print(f"  [OK] Created {total_links} hyperlinks of various types")

        # List all hyperlinks
        print("\nHyperlink Summary:")
        for i, link in enumerate(ws.hyperlinks, 1):
            link_type = link.type
            target = link.address if link.address else link.sub_address
            print(f"  {i}. {link.range}: {link_type} -> {target[:50]}...")

        # Save to file
        output_file = "outputfiles/test_hyperlinks_comprehensive.xlsx"
        print(f"\nSaving to {output_file}...")
        wb.save(output_file)
        self.assertTrue(os.path.exists(output_file))
        file_size = os.path.getsize(output_file)
        print(f"  [OK] File saved ({file_size} bytes)")

        # Load and verify all hyperlinks
        print("\nVerifying roundtrip...")
        wb_loaded = Workbook(output_file)
        ws_loaded = wb_loaded.worksheets[0]
        self.assertEqual(ws_loaded.hyperlinks.count, total_links)
        print(f"  [OK] All {ws_loaded.hyperlinks.count} hyperlinks loaded successfully")

        # Verify each type is present
        loaded_links = list(ws_loaded.hyperlinks)
        external_count = sum(1 for l in loaded_links if l.type == "External")
        internal_count = sum(1 for l in loaded_links if l.type == "Internal")
        self.assertEqual(external_count, 7, "Should have 7 external hyperlinks")
        self.assertEqual(internal_count, 2, "Should have 2 internal hyperlinks")
        print(f"  [OK] External: {external_count}, Internal: {internal_count}")

        print("\n[OK] Comprehensive mixed hyperlink test completed!")
        print("="*70 + "\n")


if __name__ == '__main__':
    # Create output directory if it doesn't exist
    os.makedirs('outputfiles', exist_ok=True)

    # Run tests
    unittest.main(verbosity=2)
