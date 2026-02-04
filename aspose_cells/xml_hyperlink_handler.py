"""
Aspose.Cells for Python - Hyperlink XML Handler Module

This module handles loading and saving hyperlinks to/from XLSX XML files
according to ECMA-376 specification.

ECMA-376 Sections:
- 18.3.1.48: hyperlink element
- 18.3.1.49: hyperlinks collection
- 12.3.2.2: Hyperlink Relationships
"""

import xml.etree.ElementTree as ET


class HyperlinkXMLLoader:
    """
    Loads hyperlinks from worksheet XML and relationship files.

    ECMA-376 compliant hyperlink loading from:
    - xl/worksheets/sheet{n}.xml (hyperlink elements)
    - xl/worksheets/_rels/sheet{n}.xml.rels (relationship targets)
    """

    def __init__(self, namespaces):
        """
        Initializes the hyperlink loader.

        Args:
            namespaces (dict): XML namespaces for parsing.
        """
        self.ns = namespaces

    def load_hyperlinks(self, worksheet, worksheet_root, zipf, sheet_num):
        """
        Loads all hyperlinks from worksheet XML and relationships.

        Args:
            worksheet: The worksheet object to load hyperlinks into.
            worksheet_root: The XML root element of the worksheet.
            zipf: ZipFile object containing the workbook data.
            sheet_num (int): Worksheet number (1-based).

        Examples:
            >>> loader.load_hyperlinks(worksheet, root, zipf, 1)
        """
        # Load relationships first (for external hyperlinks)
        relationships = self._load_relationships(zipf, sheet_num)

        # Find hyperlinks element
        hyperlinks_elem = worksheet_root.find('main:hyperlinks', namespaces=self.ns)
        if hyperlinks_elem is None:
            return  # No hyperlinks in this worksheet

        # Load each hyperlink
        for hyperlink_elem in hyperlinks_elem.findall('main:hyperlink', namespaces=self.ns):
            self._load_hyperlink(worksheet, hyperlink_elem, relationships)

    def _load_relationships(self, zipf, sheet_num):
        """
        Loads relationships from xl/worksheets/_rels/sheet{n}.xml.rels.

        Args:
            zipf: ZipFile object.
            sheet_num (int): Worksheet number (1-based).

        Returns:
            dict: Map of relationship ID to target URL.
        """
        relationships = {}
        rels_path = f'xl/worksheets/_rels/sheet{sheet_num}.xml.rels'

        try:
            rels_content = zipf.read(rels_path)
            rels_root = ET.fromstring(rels_content)

            # Namespace for relationships
            rels_ns = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}

            # Find all hyperlink relationships
            for rel_elem in rels_root.findall('rel:Relationship', namespaces=rels_ns):
                rel_type = rel_elem.get('Type', '')
                if 'hyperlink' in rel_type:
                    rel_id = rel_elem.get('Id')
                    target = rel_elem.get('Target', '')
                    relationships[rel_id] = target

        except KeyError:
            # No relationships file (no external hyperlinks)
            pass

        return relationships

    def _load_hyperlink(self, worksheet, hyperlink_elem, relationships):
        """
        Loads a single hyperlink element.

        Args:
            worksheet: The worksheet object.
            hyperlink_elem: XML element for the hyperlink.
            relationships (dict): Map of relationship IDs to targets.
        """
        # Get hyperlink attributes
        ref = hyperlink_elem.get('ref')  # Required
        location = hyperlink_elem.get('location', '')  # Internal link
        display = hyperlink_elem.get('display', '')
        tooltip = hyperlink_elem.get('tooltip', '')

        # Get relationship ID (for external links)
        r_id = hyperlink_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')

        # Determine address (external) or sub_address (internal)
        address = ""
        sub_address = ""

        if r_id and r_id in relationships:
            # External hyperlink
            address = relationships[r_id]
        elif location:
            # Internal hyperlink
            sub_address = location

        # Add hyperlink to worksheet
        try:
            hyperlink = worksheet.hyperlinks.add(
                range_address=ref,
                address=address,
                sub_address=sub_address,
                text_to_display=display,
                screen_tip=tooltip
            )
            # Store relationship ID for saving
            if r_id:
                hyperlink._relationship_id = r_id
        except ValueError as e:
            # Skip invalid hyperlinks
            print(f"Warning: Skipping invalid hyperlink at {ref}: {e}")


class HyperlinkXMLSaver:
    """
    Saves hyperlinks to worksheet XML and relationship files.

    ECMA-376 compliant hyperlink saving to:
    - xl/worksheets/sheet{n}.xml (hyperlink elements)
    - xl/worksheets/_rels/sheet{n}.xml.rels (relationship targets)
    """

    def __init__(self):
        """Initializes the hyperlink saver."""
        self._next_rel_id = 1  # Counter for relationship IDs

    def format_hyperlinks_xml(self, worksheet):
        """
        Formats hyperlinks collection as XML.

        ECMA-376 Section: 18.3.1.49 (hyperlinks collection)

        Args:
            worksheet: The worksheet containing hyperlinks.

        Returns:
            str: XML representation of hyperlinks, or empty string if none.

        Examples:
            >>> xml = saver.format_hyperlinks_xml(worksheet)
        """
        if worksheet.hyperlinks.count == 0:
            return ''

        xml = '    <hyperlinks>\n'

        for hyperlink in worksheet.hyperlinks:
            xml += self._format_hyperlink_xml(hyperlink)

        xml += '    </hyperlinks>\n'
        return xml

    def _format_hyperlink_xml(self, hyperlink):
        """
        Formats a single hyperlink as XML.

        ECMA-376 Section: 18.3.1.48 (hyperlink element)

        Args:
            hyperlink: The Hyperlink object to format.

        Returns:
            str: XML representation of the hyperlink.
        """
        attrs = [f'ref="{self._escape_xml(hyperlink.range)}"']

        # Add r:id for external links
        if hyperlink.address:
            if not hasattr(hyperlink, '_relationship_id') or not hyperlink._relationship_id:
                # Assign new relationship ID
                hyperlink._relationship_id = f'rId{self._next_rel_id}'
                self._next_rel_id += 1
            attrs.append(f'r:id="{hyperlink._relationship_id}"')

        # Add location for internal links
        if hyperlink.sub_address:
            attrs.append(f'location="{self._escape_xml(hyperlink.sub_address)}"')

        # Add display text if present
        if hyperlink.text_to_display:
            attrs.append(f'display="{self._escape_xml(hyperlink.text_to_display)}"')

        # Add tooltip if present
        if hyperlink.screen_tip:
            attrs.append(f'tooltip="{self._escape_xml(hyperlink.screen_tip)}"')

        return f'        <hyperlink {" ".join(attrs)}/>\n'

    def get_hyperlink_relationships(self, worksheet):
        """
        Gets hyperlink relationships for the worksheet.

        Returns a list of relationship entries for external hyperlinks.

        Args:
            worksheet: The worksheet containing hyperlinks.

        Returns:
            list: List of (rel_id, target) tuples for external hyperlinks.

        Examples:
            >>> rels = saver.get_hyperlink_relationships(worksheet)
            >>> for rel_id, target in rels:
            ...     print(f"{rel_id}: {target}")
        """
        relationships = []

        for hyperlink in worksheet.hyperlinks:
            if hyperlink.address:  # External hyperlink
                rel_id = hyperlink._relationship_id
                target = hyperlink.address
                relationships.append((rel_id, target))

        return relationships

    def reset_relationship_counter(self):
        """
        Resets the relationship ID counter.

        Should be called before processing each worksheet.
        """
        self._next_rel_id = 1

    def _escape_xml(self, text):
        """
        Escapes special characters for XML.

        Args:
            text (str): Text to escape.

        Returns:
            str: XML-escaped text.
        """
        if not text:
            return ""
        text = str(text)
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        text = text.replace('"', '&quot;')
        text = text.replace("'", '&apos;')
        return text


class HyperlinkRelationshipWriter:
    """
    Writes hyperlink relationships to _rels files.

    Creates or updates xl/worksheets/_rels/sheet{n}.xml.rels with hyperlink relationships.
    """

    @staticmethod
    def format_relationships_xml(relationships, existing_rels=None):
        """
        Formats relationships as XML.

        Args:
            relationships (list): List of (rel_id, target) tuples for hyperlinks.
            existing_rels (list, optional): Existing non-hyperlink relationships to preserve.

        Returns:
            str: Complete relationships XML content.

        Examples:
            >>> xml = writer.format_relationships_xml([('rId1', 'https://example.com')])
        """
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'

        # Add existing non-hyperlink relationships first
        if existing_rels:
            for rel_id, rel_type, target, target_mode in existing_rels:
                mode_attr = f' TargetMode="{target_mode}"' if target_mode else ''
                content += f'    <Relationship Id="{rel_id}" Type="{rel_type}" Target="{target}"{mode_attr}/>\n'

        # Add hyperlink relationships
        for rel_id, target in relationships:
            escaped_target = HyperlinkRelationshipWriter._escape_xml(target)
            content += f'    <Relationship Id="{rel_id}" '
            content += f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            content += f'Target="{escaped_target}" TargetMode="External"/>\n'

        content += '</Relationships>\n'
        return content

    @staticmethod
    def _escape_xml(text):
        """Escapes special characters for XML."""
        if not text:
            return ""
        text = str(text)
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        text = text.replace('"', '&quot;')
        text = text.replace("'", '&apos;')
        return text
