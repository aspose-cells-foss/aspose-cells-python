import xml.etree.ElementTree as ET

class SharedStringTable:
    """Manages the Shared String Table for XLSX files according to ECMA-376 specification."""
    
    def __init__(self):
        """Initialize an empty shared string table."""
        self.strings = []
        self.string_to_index = {}
        self.count = 0  # Track total occurrences of strings
    
    def add_string(self, text):
        """
        Add a string to the shared string table and return its index.
        If the string already exists, return the existing index.
        
        Args:
            text (str): The string to add to the table
            
        Returns:
            int: The index of the string in the shared string table
        """
        if text is None:
            return None
            
        # Increment total count for ECMA-376 compliance
        self.count += 1
            
        if text in self.string_to_index:
            return self.string_to_index[text]
        
        index = len(self.strings)
        self.strings.append(text)
        self.string_to_index[text] = index
        return index
    
    def get_string(self, index):
        """
        Get a string from the shared string table by its index.
        
        Args:
            index (int): The index of the string to retrieve
            
        Returns:
            str: The string at the given index, or None if index is invalid
        """
        if index is None or index < 0 or index >= len(self.strings):
            return None
        return self.strings[index]
    
    def to_xml(self):
        """
        Convert the shared string table to XML format for XLSX files.
        
        Returns:
            str: XML representation of the shared string table
        """
        root = ET.Element('sst', {
            'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'count': str(self.count),  # Total occurrences
            'uniqueCount': str(len(self.strings))  # Unique strings
        })
        
        for text in self.strings:
            si = ET.SubElement(root, 'si')
            t = ET.SubElement(si, 't')
            t.text = text
        
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(root, encoding='unicode')
    
    @classmethod
    def from_xml(cls, xml_content):
        """
        Create a SharedStringTable from XML content.
        
        Args:
            xml_content (str): XML content of the shared string table
            
        Returns:
            SharedStringTable: A new SharedStringTable instance
        """
        sst = cls()
        if not xml_content:
            return sst
            
        root = ET.fromstring(xml_content)
        ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        
        for si in root.findall('.//ns:si', ns):
            text_parts = [
                t.text if t.text is not None else ''
                for t in si.findall('.//ns:t', ns)
            ]
            # Preserve indices for rich text and empty strings
            sst.add_string(''.join(text_parts))
        
        return sst
    
    def __len__(self):
        """Return the number of strings in the table."""
        return len(self.strings)
