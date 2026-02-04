"""
Test cases for workbook document properties persistence.

Tests roundtrip persistence of document properties according to ECMA-376 specification.
Document properties are stored in docProps/core.xml and docProps/app.xml files.
"""

import os
import sys
import unittest
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from aspose_cells import Workbook


class TestDocumentProperties(unittest.TestCase):
    """Test cases for document properties persistence."""

    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputfiles')
        os.makedirs(self.test_dir, exist_ok=True)
        self.ns = {
            'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
            'dc': 'http://purl.org/dc/elements/1.1/',
            'dcterms': 'http://purl.org/dc/terms/'
        }

    def test_core_properties_basic(self):
        """Test that basic core properties persist correctly."""
        wb = Workbook()
        
        # Set core properties
        wb.document_properties.core.title = "Test Document"
        wb.document_properties.core.subject = "Test Subject"
        wb.document_properties.core.creator = "Test Author"
        wb.document_properties.core.keywords = "test, keywords"
        wb.document_properties.core.description = "Test description"

        # Save and reload
        path = os.path.join(self.test_dir, "core_properties_basic.xlsx")
        wb.save(path)

        # Verify XML structure
        with zipfile.ZipFile(path, 'r') as zf:
            core_xml = zf.read('docProps/core.xml').decode('utf-8')
            root = ET.fromstring(core_xml)
            
            title = root.find('dc:title', self.ns)
            self.assertIsNotNone(title)
            self.assertEqual(title.text, "Test Document")
            
            subject = root.find('dc:subject', self.ns)
            self.assertIsNotNone(subject)
            self.assertEqual(subject.text, "Test Subject")
            
            creator = root.find('dc:creator', self.ns)
            self.assertIsNotNone(creator)
            self.assertEqual(creator.text, "Test Author")
            
            keywords = root.find('cp:keywords', self.ns)
            self.assertIsNotNone(keywords)
            self.assertEqual(keywords.text, "test, keywords")
            
            description = root.find('dc:description', self.ns)
            self.assertIsNotNone(description)
            self.assertEqual(description.text, "Test description")

        # Reload and verify
        wb2 = Workbook(path)
        self.assertEqual(wb2.document_properties.core.title, "Test Document")
        self.assertEqual(wb2.document_properties.core.subject, "Test Subject")
        self.assertEqual(wb2.document_properties.core.creator, "Test Author")
        self.assertEqual(wb2.document_properties.core.keywords, "test, keywords")
        self.assertEqual(wb2.document_properties.core.description, "Test description")

    def test_core_properties_extended(self):
        """Test that extended core properties persist correctly."""
        wb = Workbook()
        
        # Set extended core properties
        wb.document_properties.core.last_modified_by = "Second Author"
        wb.document_properties.core.revision = "2"
        wb.document_properties.core.category = "Reports"
        wb.document_properties.core.content_status = "Draft"

        # Save and reload
        path = os.path.join(self.test_dir, "core_properties_extended.xlsx")
        wb.save(path)

        # Verify XML structure
        with zipfile.ZipFile(path, 'r') as zf:
            core_xml = zf.read('docProps/core.xml').decode('utf-8')
            root = ET.fromstring(core_xml)
            
            last_modified_by = root.find('cp:lastModifiedBy', self.ns)
            self.assertIsNotNone(last_modified_by)
            self.assertEqual(last_modified_by.text, "Second Author")
            
            revision = root.find('cp:revision', self.ns)
            self.assertIsNotNone(revision)
            self.assertEqual(revision.text, "2")
            
            category = root.find('cp:category', self.ns)
            self.assertIsNotNone(category)
            self.assertEqual(category.text, "Reports")
            
            content_status = root.find('cp:contentStatus', self.ns)
            self.assertIsNotNone(content_status)
            self.assertEqual(content_status.text, "Draft")

        # Reload and verify
        wb2 = Workbook(path)
        self.assertEqual(wb2.document_properties.core.last_modified_by, "Second Author")
        self.assertEqual(wb2.document_properties.core.revision, "2")
        self.assertEqual(wb2.document_properties.core.category, "Reports")
        self.assertEqual(wb2.document_properties.core.content_status, "Draft")

    def test_core_properties_dates(self):
        """Test that date properties persist correctly."""
        wb = Workbook()
        
        # Set date properties
        test_created = datetime(2024, 1, 15, 10, 30)
        test_modified = datetime(2024, 1, 20, 14, 45, 0)
        wb.document_properties.core.created = test_created
        wb.document_properties.core.modified = test_modified

        # Save and reload
        path = os.path.join(self.test_dir, "core_properties_dates.xlsx")
        wb.save(path)

        # Verify XML structure
        with zipfile.ZipFile(path, 'r') as zf:
            core_xml = zf.read('docProps/core.xml').decode('utf-8')
            root = ET.fromstring(core_xml)
            
            created = root.find('dcterms:created', self.ns)
            self.assertIsNotNone(created)
            self.assertIn('2024-01-15T10:30:00', created.text)
            
            modified = root.find('dcterms:modified', self.ns)
            self.assertIsNotNone(modified)
            self.assertIn('2024-01-20T14:45:00', modified.text)

        # Reload and verify
        wb2 = Workbook(path)
        self.assertIsNotNone(wb2.document_properties.core.created)
        self.assertIsNotNone(wb2.document_properties.core.modified)

    def test_extended_properties_basic(self):
        """Test that basic extended properties persist correctly."""
        wb = Workbook()
        
        # Set extended properties
        wb.document_properties.extended.application = "Aspose.Cells for Python"
        wb.document_properties.extended.app_version = "16.0.0"
        wb.document_properties.extended.company = "Test Company"
        wb.document_properties.extended.manager = "Test Manager"

        # Save and reload
        path = os.path.join(self.test_dir, "extended_properties_basic.xlsx")
        wb.save(path)

        # Verify XML structure
        with zipfile.ZipFile(path, 'r') as zf:
            app_xml = zf.read('docProps/app.xml').decode('utf-8')
            root = ET.fromstring(app_xml)
            
            application = root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Application')
            self.assertIsNotNone(application)
            self.assertEqual(application.text, "Aspose.Cells for Python")
            
            app_version = root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}AppVersion')
            self.assertIsNotNone(app_version)
            self.assertEqual(app_version.text, "16.0.0")
            
            company = root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Company')
            self.assertIsNotNone(company)
            self.assertEqual(company.text, "Test Company")
            
            manager = root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Manager')
            self.assertIsNotNone(manager)
            self.assertEqual(manager.text, "Test Manager")

        # Reload and verify
        wb2 = Workbook(path)
        self.assertEqual(wb2.document_properties.extended.application, "Aspose.Cells for Python")
        self.assertEqual(wb2.document_properties.extended.app_version, "16.0.0")
        self.assertEqual(wb2.document_properties.extended.company, "Test Company")
        self.assertEqual(wb2.document_properties.extended.manager, "Test Manager")

    def test_extended_properties_flags(self):
        """Test that extended property flags persist correctly."""
        wb = Workbook()
        
        # Set extended property flags
        wb.document_properties.extended.scale_crop = True
        wb.document_properties.extended.links_up_to_date = True
        wb.document_properties.extended.shared_doc = True
        wb.document_properties.extended.doc_security = 1

        # Save and reload
        path = os.path.join(self.test_dir, "extended_properties_flags.xlsx")
        wb.save(path)

        # Verify XML structure
        with zipfile.ZipFile(path, 'r') as zf:
            app_xml = zf.read('docProps/app.xml').decode('utf-8')
            root = ET.fromstring(app_xml)
            
            scale_crop = root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}ScaleCrop')
            self.assertIsNotNone(scale_crop)
            self.assertEqual(scale_crop.text, "true")
            
            links_up_to_date = root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}LinksUpToDate')
            self.assertIsNotNone(links_up_to_date)
            self.assertEqual(links_up_to_date.text, "true")
            
            shared_doc = root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}SharedDoc')
            self.assertIsNotNone(shared_doc)
            self.assertEqual(shared_doc.text, "true")
            
            doc_security = root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}DocSecurity')
            self.assertIsNotNone(doc_security)
            self.assertEqual(doc_security.text, "1")

        # Reload and verify
        wb2 = Workbook(path)
        self.assertTrue(wb2.document_properties.extended.scale_crop)
        self.assertTrue(wb2.document_properties.extended.links_up_to_date)
        self.assertTrue(wb2.document_properties.extended.shared_doc)
        self.assertEqual(wb2.document_properties.extended.doc_security, 1)

    def test_extended_properties_hyperlink_base(self):
        """Test that hyperlink base property persists correctly."""
        wb = Workbook()
        
        # Set hyperlink base
        wb.document_properties.extended.hyperlink_base = "https://example.com/docs/"

        # Save and reload
        path = os.path.join(self.test_dir, "extended_properties_hyperlink.xlsx")
        wb.save(path)

        # Verify XML structure
        with zipfile.ZipFile(path, 'r') as zf:
            app_xml = zf.read('docProps/app.xml').decode('utf-8')
            root = ET.fromstring(app_xml)
            
            hyperlink_base = root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}HyperlinkBase')
            self.assertIsNotNone(hyperlink_base)
            self.assertEqual(hyperlink_base.text, "https://example.com/docs/")

        # Reload and verify
        wb2 = Workbook(path)
        self.assertEqual(wb2.document_properties.extended.hyperlink_base, "https://example.com/docs/")

    def test_convenience_properties(self):
        """Test that convenience properties work correctly."""
        wb = Workbook()
        
        # Set convenience properties
        wb.document_properties.title = "Convenience Title"
        wb.document_properties.subject = "Convenience Subject"
        wb.document_properties.author = "Convenience Author"
        wb.document_properties.creator = "Creator Name"
        wb.document_properties.keywords = "convenience, test"
        wb.document_properties.comments = "Convenience comments"
        wb.document_properties.category = "Convenience Category"
        wb.document_properties.company = "Convenience Company"
        wb.document_properties.manager = "Convenience Manager"

        # Save and reload
        path = os.path.join(self.test_dir, "convenience_properties.xlsx")
        wb.save(path)

        # Reload and verify
        wb2 = Workbook(path)
        self.assertEqual(wb2.document_properties.title, "Convenience Title")
        self.assertEqual(wb2.document_properties.subject, "Convenience Subject")
        self.assertEqual(wb2.document_properties.author, "Creator Name")
        self.assertEqual(wb2.document_properties.creator, "Creator Name")
        self.assertEqual(wb2.document_properties.keywords, "convenience, test")
        self.assertEqual(wb2.document_properties.comments, "Convenience comments")
        self.assertEqual(wb2.document_properties.category, "Convenience Category")
        self.assertEqual(wb2.document_properties.company, "Convenience Company")
        self.assertEqual(wb2.document_properties.manager, "Convenience Manager")

    def test_comprehensive_document_properties(self):
        """Test comprehensive roundtrip of all document properties."""
        wb = Workbook()
        
        # Set all core properties
        wb.document_properties.core.title = "Comprehensive Test"
        wb.document_properties.core.subject = "Testing All Properties"
        wb.document_properties.core.creator = "Test Creator"
        wb.document_properties.core.keywords = "comprehensive, test, all"
        wb.document_properties.core.description = "Testing all document properties"
        wb.document_properties.core.last_modified_by = "Second Creator"
        wb.document_properties.core.revision = "3"
        wb.document_properties.core.category = "Test Reports"
        wb.document_properties.core.content_status = "Final"

        # Set all extended properties
        wb.document_properties.extended.application = "Test Application"
        wb.document_properties.extended.app_version = "1.0.0"
        wb.document_properties.extended.company = "Test Company Inc."
        wb.document_properties.extended.manager = "Test Manager"
        wb.document_properties.extended.hyperlink_base = "https://test.com/"
        wb.document_properties.extended.scale_crop = False
        wb.document_properties.extended.links_up_to_date = False
        wb.document_properties.extended.shared_doc = False
        wb.document_properties.extended.doc_security = 0

        # Add some data
        ws = wb.worksheets[0]
        ws.cells["A1"].value = "Test"
        ws.cells["A2"].value = 123

        # Save and reload
        path = os.path.join(self.test_dir, "comprehensive_document_properties.xlsx")
        wb.save(path)

        # Reload and verify
        wb2 = Workbook(path)
        ws2 = wb2.worksheets[0]

        # Verify core properties
        self.assertEqual(wb2.document_properties.core.title, "Comprehensive Test")
        self.assertEqual(wb2.document_properties.core.subject, "Testing All Properties")
        self.assertEqual(wb2.document_properties.core.creator, "Test Creator")
        self.assertEqual(wb2.document_properties.core.keywords, "comprehensive, test, all")
        self.assertEqual(wb2.document_properties.core.description, "Testing all document properties")
        self.assertEqual(wb2.document_properties.core.last_modified_by, "Second Creator")
        self.assertEqual(wb2.document_properties.core.revision, "3")
        self.assertEqual(wb2.document_properties.core.category, "Test Reports")
        self.assertEqual(wb2.document_properties.core.content_status, "Final")

        # Verify extended properties
        self.assertEqual(wb2.document_properties.extended.application, "Test Application")
        self.assertEqual(wb2.document_properties.extended.app_version, "1.0.0")
        self.assertEqual(wb2.document_properties.extended.company, "Test Company Inc.")
        self.assertEqual(wb2.document_properties.extended.manager, "Test Manager")
        self.assertEqual(wb2.document_properties.extended.hyperlink_base, "https://test.com/")
        self.assertFalse(wb2.document_properties.extended.scale_crop)
        self.assertFalse(wb2.document_properties.extended.links_up_to_date)
        self.assertFalse(wb2.document_properties.extended.shared_doc)
        self.assertEqual(wb2.document_properties.extended.doc_security, 0)

        # Verify data
        self.assertEqual(ws2.cells["A1"].value, "Test")
        self.assertEqual(ws2.cells["A2"].value, 123)

    def test_document_properties_with_special_characters(self):
        """Test that special characters in properties are handled correctly."""
        wb = Workbook()
        
        # Set properties with special characters
        wb.document_properties.core.title = "Test <Special> & Characters"
        wb.document_properties.core.description = "Description with 'quotes' and \"double quotes\""
        wb.document_properties.core.keywords = "test & special, <chars>"

        # Save and reload
        path = os.path.join(self.test_dir, "special_characters_properties.xlsx")
        wb.save(path)

        # Reload and verify
        wb2 = Workbook(path)
        self.assertEqual(wb2.document_properties.core.title, "Test <Special> & Characters")
        self.assertEqual(wb2.document_properties.core.description, "Description with 'quotes' and \"double quotes\"")
        self.assertEqual(wb2.document_properties.core.keywords, "test & special, <chars>")

    def test_document_properties_with_unicode(self):
        """Test that Unicode characters in properties are handled correctly."""
        wb = Workbook()
        
        # Set properties with Unicode characters
        wb.document_properties.core.title = "测试文档"
        wb.document_properties.core.creator = "作者"
        wb.document_properties.core.description = "这是一份测试文档，包含Unicode字符"

        # Save and reload
        path = os.path.join(self.test_dir, "unicode_properties.xlsx")
        wb.save(path)

        # Reload and verify
        wb2 = Workbook(path)
        self.assertEqual(wb2.document_properties.core.title, "测试文档")
        self.assertEqual(wb2.document_properties.core.creator, "作者")
        self.assertEqual(wb2.document_properties.core.description, "这是一份测试文档，包含Unicode字符")

    def test_document_properties_default_values(self):
        """Test that default properties are set correctly when not explicitly set."""
        wb = Workbook()
        
        # Add some data
        ws = wb.worksheets[0]
        ws.cells["A1"].value = "Test"

        # Save and reload
        path = os.path.join(self.test_dir, "default_properties.xlsx")
        wb.save(path)

        # Reload and verify defaults
        wb2 = Workbook(path)
        
        # Core properties should be None or have default values
        self.assertIsNone(wb2.document_properties.core.title)
        self.assertIsNone(wb2.document_properties.core.subject)
        self.assertIsNone(wb2.document_properties.core.creator)
        
        # Extended properties should have defaults
        self.assertEqual(wb2.document_properties.extended.application, "Microsoft Excel")
        self.assertEqual(wb2.document_properties.extended.doc_security, 0)
        self.assertFalse(wb2.document_properties.extended.scale_crop)
        self.assertFalse(wb2.document_properties.extended.links_up_to_date)
        self.assertFalse(wb2.document_properties.extended.shared_doc)

    def test_document_properties_multiple_workbooks(self):
        """Test that properties work correctly with multiple workbooks."""
        # Create first workbook with properties
        wb1 = Workbook()
        wb1.document_properties.core.title = "Workbook 1"
        wb1.document_properties.core.creator = "Author 1"
        wb1.document_properties.extended.company = "Company 1"
        path1 = os.path.join(self.test_dir, "workbook1_properties.xlsx")
        wb1.save(path1)

        # Create second workbook with different properties
        wb2 = Workbook()
        wb2.document_properties.core.title = "Workbook 2"
        wb2.document_properties.core.creator = "Author 2"
        wb2.document_properties.extended.company = "Company 2"
        path2 = os.path.join(self.test_dir, "workbook2_properties.xlsx")
        wb2.save(path2)

        # Verify first workbook
        wb1_reloaded = Workbook(path1)
        self.assertEqual(wb1_reloaded.document_properties.core.title, "Workbook 1")
        self.assertEqual(wb1_reloaded.document_properties.core.creator, "Author 1")
        self.assertEqual(wb1_reloaded.document_properties.extended.company, "Company 1")

        # Verify second workbook
        wb2_reloaded = Workbook(path2)
        self.assertEqual(wb2_reloaded.document_properties.core.title, "Workbook 2")
        self.assertEqual(wb2_reloaded.document_properties.core.creator, "Author 2")
        self.assertEqual(wb2_reloaded.document_properties.extended.company, "Company 2")

    def test_document_properties_xml_structure(self):
        """Test that the XML structure of document properties is correct."""
        wb = Workbook()
        
        # Set some properties
        wb.document_properties.core.title = "XML Structure Test"
        wb.document_properties.extended.company = "Test Company"

        # Save
        path = os.path.join(self.test_dir, "xml_structure_properties.xlsx")
        wb.save(path)

        # Verify core.xml structure
        with zipfile.ZipFile(path, 'r') as zf:
            core_xml = zf.read('docProps/core.xml').decode('utf-8')
            self.assertIn('cp:coreProperties', core_xml)
            self.assertIn('xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"', core_xml)
            self.assertIn('xmlns:dc="http://purl.org/dc/elements/1.1/"', core_xml)
            self.assertIn('xmlns:dcterms="http://purl.org/dc/terms/"', core_xml)

        # Verify app.xml structure
        with zipfile.ZipFile(path, 'r') as zf:
            app_xml = zf.read('docProps/app.xml').decode('utf-8')
            self.assertIn('Properties', app_xml)
            self.assertIn('xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"', app_xml)
            self.assertIn('xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"', app_xml)


if __name__ == '__main__':
    unittest.main()
