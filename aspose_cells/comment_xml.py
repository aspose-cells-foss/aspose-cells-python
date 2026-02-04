"""
Aspose.Cells for Python - Comment XML Module

This module handles reading and writing comment data to/from XML format
according to ECMA-376 standards.

Comment data is stored in:
- xl/comments{n}.xml - Comment text and authors
- xl/drawings/vmlDrawing{n}.vml - Comment positioning and sizing (VML format)
"""

import re
import xml.etree.ElementTree as ET


# Default comment dimensions (Excel standard)
DEFAULT_COMMENT_WIDTH = 96  # points
DEFAULT_COMMENT_HEIGHT = 55.5  # points

# Standard Excel dimensions for anchor calculations
AVG_COL_WIDTH_PT = 48  # Standard column width in points
AVG_ROW_HEIGHT_PT = 15  # Standard row height in points


class CommentXMLWriter:
    """
    Handles writing comment data to XML format.

    Writes both comments XML and VML drawing files for proper
    Excel compatibility according to ECMA-376.
    """

    def __init__(self):
        """Initialize the CommentXMLWriter."""
        pass

    @staticmethod
    def escape_xml(text):
        """
        Escapes special XML characters in text.

        Args:
            text: The text to escape.

        Returns:
            str: The escaped text.
        """
        if text is None:
            return ''
        text = str(text)
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        text = text.replace('"', '&quot;')
        text = text.replace("'", '&apos;')
        return text

    @staticmethod
    def worksheet_has_comments(worksheet):
        """
        Checks if a worksheet has any comments.

        Args:
            worksheet: The worksheet to check.

        Returns:
            bool: True if the worksheet has comments, False otherwise.
        """
        for cell in worksheet.cells._cells.values():
            if cell.has_comment():
                return True
        return False

    @staticmethod
    def cell_reference_sort_key(ref):
        """
        Converts a cell reference to a (row, col) tuple for sorting.

        Args:
            ref: Cell reference string (e.g., 'A1', 'BC23').

        Returns:
            tuple: (row, col) where both are 1-based integers.
        """
        col_str = ''
        row_str = ''
        for char in ref:
            if char.isalpha():
                col_str += char
            else:
                row_str += char

        # Convert column letters to number
        col = 0
        for char in col_str.upper():
            col = col * 26 + (ord(char) - ord('A') + 1)

        row = int(row_str) if row_str else 0
        return (row, col)

    def write_comments_xml(self, zipf, worksheet, sheet_num):
        """
        Writes xl/comments{sheet_num}.xml file for a worksheet.

        According to ECMA-376 Part 1, Section 18.7.3, comments are stored
        in a separate XML file with authors and comment text.

        ECMA-376 compliant format includes:
        - Rich text format with <r> (run) elements
        - Author name prepended to comment text with bold formatting
        - Font properties (size, color, font family)
        - shapeId attribute for each comment

        Args:
            zipf: The ZIP file object to write to.
            worksheet: The worksheet object.
            sheet_num: The worksheet number (1-based).
        """
        if not self.worksheet_has_comments(worksheet):
            return

        # Collect all unique authors and comments
        authors = set()
        comments_data = []

        for ref, cell in worksheet.cells._cells.items():
            if cell.has_comment():
                comment = cell.get_comment()
                author = comment.get('author', '')
                text = comment.get('text', '')

                if author not in authors:
                    authors.add(author)

                comments_data.append({
                    'ref': ref,
                    'author': author,
                    'text': text
                })

        authors_list = list(authors)

        # Build comments XML with ECMA-376 compliant rich text format
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        content += 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        content += 'mc:Ignorable="xr" '
        content += 'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">\n'

        # Write authors section (no count attribute per ECMA-376)
        content += '    <authors>\n'
        for author in authors_list:
            escaped_author = self.escape_xml(author)
            content += f'        <author>{escaped_author}</author>\n'
        content += '    </authors>\n'

        # Write comment list with rich text format (no count attribute per ECMA-376)
        content += '    <commentList>\n'
        for comment_data in comments_data:
            author_idx = authors_list.index(comment_data['author'])
            author_name = comment_data['author']
            comment_text = comment_data['text']

            # shapeId is always 0 per Excel's implementation
            content += f'        <comment ref="{comment_data["ref"]}" authorId="{author_idx}" shapeId="0">\n'
            content += '            <text>\n'

            # Author name run (bold with formatting)
            content += '                <r>\n'
            content += '                    <rPr>\n'
            content += '                        <b/>\n'
            content += '                        <sz val="9"/>\n'
            content += '                        <color indexed="81"/>\n'
            content += '                        <rFont val="Tahoma"/>\n'
            content += '                        <family val="2"/>\n'
            content += '                    </rPr>\n'
            content += f'                    <t>{self.escape_xml(author_name)}:</t>\n'
            content += '                </r>\n'

            # Comment text run
            content += '                <r>\n'
            content += '                    <rPr>\n'
            content += '                        <sz val="9"/>\n'
            content += '                        <color indexed="81"/>\n'
            content += '                        <rFont val="Tahoma"/>\n'
            content += '                        <family val="2"/>\n'
            content += '                    </rPr>\n'
            content += f'                    <t xml:space="preserve">{self.escape_xml(comment_text)}</t>\n'
            content += '                </r>\n'

            content += '            </text>\n'
            content += '        </comment>\n'
        content += '    </commentList>\n'

        content += '</comments>\n'
        zipf.writestr(f'xl/comments{sheet_num}.xml', content)

    def write_vml_drawing_xml(self, zipf, worksheet, sheet_num):
        """
        Writes VML drawing XML for comment positioning.

        According to ECMA-376 Part 1, Section 18.3.1.43, legacy drawings
        are used for backward compatibility with older Excel versions and
        contain shape information for comment positioning.

        The Anchor element defines cell-relative position and size using 8 values:
        colStart, xOffset, rowStart, yOffset, colEnd, xOffsetEnd, rowEnd, yOffsetEnd

        Args:
            zipf: The ZIP file object to write to.
            worksheet: The worksheet object.
            sheet_num: The worksheet number (1-based).
        """
        if not self.worksheet_has_comments(worksheet):
            return

        content = '<xml xmlns:v="urn:schemas-microsoft-com:vml"\n'
        content += '     xmlns:o="urn:schemas-microsoft-com:office:office"\n'
        content += '     xmlns:x="urn:schemas-microsoft-com:office:excel">\n'
        content += ' <o:shapelayout v:ext="edit">\n'
        content += '  <o:idmap v:ext="edit" data="1"/>\n'
        content += ' </o:shapelayout>\n'
        content += ' <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"\n'
        content += '              path="m,l,21600r21600,l21600,xe">\n'
        content += '  <v:stroke joinstyle="miter"/>\n'
        content += '  <v:path gradientshapeok="t" o:connecttype="rect"/>\n'
        content += ' </v:shapetype>\n'

        # Add shapes for each comment
        shape_id = 1025
        for ref, cell in worksheet.cells._cells.items():
            if cell.has_comment():
                row, col = self.cell_reference_sort_key(ref)
                comment = cell.get_comment()

                # Get comment size (default to Excel's default size if not specified)
                width = comment.get('width')
                height = comment.get('height')

                # Use Excel defaults if not specified
                if width is None:
                    width = DEFAULT_COMMENT_WIDTH
                if height is None:
                    height = DEFAULT_COMMENT_HEIGHT

                # Calculate anchor values from width/height
                anchor = self._calculate_anchor(row, col, width, height)

                content += f' <v:shape id="_x0000_s{shape_id}" type="#_x0000_t202"\n'
                content += f'          style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:{width}pt;height:{height}pt;z-index:{shape_id-1024};visibility:hidden"\n'
                content += '          fillcolor="infoBackground [80]" strokecolor="none [81]"\n'
                content += '          o:insetmode="auto">\n'
                content += '  <v:fill color2="infoBackground [80]"/>\n'
                content += '  <v:shadow color="none [81]"/>\n'
                content += '  <v:textbox>\n'
                content += '   <div style="text-align:left"></div>\n'
                content += '  </v:textbox>\n'
                content += '  <x:ClientData ObjectType="Note">\n'
                content += '   <x:MoveWithCells/>\n'
                content += '   <x:SizeWithCells/>\n'
                content += f'   <x:Anchor>{anchor}</x:Anchor>\n'
                content += f'   <x:Row>{row-1}</x:Row>\n'
                content += f'   <x:Column>{col-1}</x:Column>\n'
                content += '  </x:ClientData>\n'
                content += ' </v:shape>\n'

                shape_id += 1

        content += '</xml>\n'
        zipf.writestr(f'xl/drawings/vmlDrawing{sheet_num}.vml', content)

    def _calculate_anchor(self, row, col, width, height):
        """
        Calculate anchor coordinates from width/height in points.

        Anchor format: colStart, xOffset, rowStart, yOffset, colEnd, xOffsetEnd, rowEnd, yOffsetEnd

        Args:
            row: Cell row (1-based).
            col: Cell column (1-based).
            width: Comment width in points.
            height: Comment height in points.

        Returns:
            str: Anchor string with 8 comma-separated values.
        """
        # Calculate how many columns/rows the comment spans
        col_span = int(width / AVG_COL_WIDTH_PT)
        row_span = int(height / AVG_ROW_HEIGHT_PT)

        # Calculate offsets (remainder after full column/row spans)
        # X offsets in 1/256ths of column width
        x_offset_start = 12  # Small offset to right of cell
        x_offset_end = int(((width % AVG_COL_WIDTH_PT) / AVG_COL_WIDTH_PT) * 256)

        # Y offsets in points or 1/256ths of row height
        y_offset_start = 4  # Small offset below row
        y_offset_end = int(((height % AVG_ROW_HEIGHT_PT) / AVG_ROW_HEIGHT_PT) * 256)

        # Position comment starting from cell position
        # Excel typically positions comments 2 rows above and at the same column
        anchor_col_start = col - 1  # 0-based
        anchor_row_start = max(0, row - 3)  # Start 2 rows above (0-based, so -3)
        anchor_col_end = anchor_col_start + col_span
        anchor_row_end = anchor_row_start + row_span

        return f"{anchor_col_start}, {x_offset_start}, {anchor_row_start}, {y_offset_start}, {anchor_col_end}, {x_offset_end}, {anchor_row_end}, {y_offset_end}"


class CommentXMLReader:
    """
    Handles reading comment data from XML format.

    Reads both comments XML and VML drawing files to restore
    comment text, authors, and sizing information.
    """

    # XML namespaces for parsing
    NS = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    def __init__(self):
        """Initialize the CommentXMLReader."""
        pass

    def load_comments(self, zipf, worksheet, sheet_num):
        """
        Loads comments from the comments XML file for a worksheet.

        According to ECMA-376 Part 1, Section 18.7.3, comments are stored in
        xl/comments{i}.xml and contain author information and comment text.

        ECMA-376 compliant format includes rich text with <r> (run) elements.
        The author name is typically prepended to the comment text with bold formatting.

        Args:
            zipf: A ZipFile object containing the workbook data.
            worksheet: The worksheet object to load comments into.
            sheet_num: The worksheet number (1-based).
        """
        try:
            comments_content = zipf.read(f'xl/comments{sheet_num}.xml')
            comments_root = ET.fromstring(comments_content)

            # Get authors list (ECMA-376: authors are stored separately)
            authors = []
            authors_elem = comments_root.find('.//main:authors', namespaces=self.NS)
            if authors_elem is not None:
                for author in authors_elem.findall('main:author', namespaces=self.NS):
                    authors.append(author.text if author.text else '')

            # Load comments
            comments_elem = comments_root.find('.//main:commentList', namespaces=self.NS)
            if comments_elem is not None:
                for comment in comments_elem.findall('main:comment', namespaces=self.NS):
                    # Get cell reference
                    cell_ref = comment.get('ref')

                    # Get author index
                    author_idx = int(comment.get('authorId', 0))
                    author = authors[author_idx] if author_idx < len(authors) else ''

                    # Get comment text from rich text runs
                    # ECMA-376: text can be in <t> elements within <r> (run) elements
                    comment_text = ''
                    text_elem = comment.find('.//main:text', namespaces=self.NS)
                    if text_elem is not None:
                        # Collect all text from <t> elements
                        for t in text_elem.findall('.//main:t', namespaces=self.NS):
                            if t.text:
                                comment_text += t.text

                    # Remove author prefix if present (e.g., "Author Name:")
                    # Excel formats comments with author name prepended
                    if author and comment_text.startswith(f'{author}:'):
                        comment_text = comment_text[len(author)+1:].lstrip()

                    # Create cell if it doesn't exist (comments can exist on empty cells)
                    if cell_ref not in worksheet.cells._cells:
                        from .cell import Cell
                        worksheet.cells._cells[cell_ref] = Cell(None, None)

                    # Set comment on the cell
                    cell = worksheet.cells._cells[cell_ref]
                    cell.set_comment(comment_text, author)

            # Load VML drawing for comment sizes
            self.load_vml_drawing(zipf, worksheet, sheet_num)
        except KeyError:
            # Comments file not found, skip
            pass

    def load_vml_drawing(self, zipf, worksheet, sheet_num):
        """
        Loads VML drawing for comment positioning and sizing.

        Parses the Anchor element to extract comment size information
        and associates it with the corresponding cells.

        Args:
            zipf: A ZipFile object containing the workbook data.
            worksheet: The worksheet object.
            sheet_num: The worksheet number (1-based).
        """
        try:
            vml_content = zipf.read(f'xl/drawings/vmlDrawing{sheet_num}.vml').decode('utf-8')

            # Split by v:shape tags to process each shape individually
            # Match only actual comment shapes (with id="_x0000_s..."), not shapetype
            shapes = re.findall(r'<v:shape[^>]*id="_x0000_s\d+"[^>]*>.*?</v:shape>', vml_content, re.DOTALL)

            for shape in shapes:
                # Extract anchor, row, and column
                anchor_match = re.search(r'<x:Anchor>([^<]+)</x:Anchor>', shape)
                row_match = re.search(r'<x:Row>(\d+)</x:Row>', shape)
                col_match = re.search(r'<x:Column>(\d+)</x:Column>', shape)

                if anchor_match and row_match and col_match:
                    anchor_str = anchor_match.group(1).strip()
                    row = int(row_match.group(1)) + 1  # VML uses 0-based indexing
                    col = int(col_match.group(1)) + 1

                    # Parse anchor values and calculate size
                    width, height = self._parse_anchor_to_size(anchor_str)

                    if width is not None and height is not None:
                        # Find the cell with this comment
                        from .cells import Cells
                        col_letter = Cells.column_letter_from_index(col)
                        cell_ref = f"{col_letter}{row}"

                        if cell_ref in worksheet.cells._cells:
                            cell = worksheet.cells._cells[cell_ref]
                            if cell.has_comment():
                                cell._comment['width'] = round(width, 1)
                                cell._comment['height'] = round(height, 1)

        except (KeyError, Exception):
            # VML drawing not found or parsing error, skip
            pass

    def _parse_anchor_to_size(self, anchor_str):
        """
        Parse anchor string and calculate width/height in points.

        Anchor format: colStart, xOffset, rowStart, yOffset, colEnd, xOffsetEnd, rowEnd, yOffsetEnd

        Args:
            anchor_str: The anchor string with 8 comma-separated values.

        Returns:
            tuple: (width, height) in points, or (None, None) if parsing fails.
        """
        try:
            anchor_values = [x.strip() for x in anchor_str.split(',')]
            if len(anchor_values) == 8:
                col_start = int(anchor_values[0])
                col_end = int(anchor_values[4])
                x_offset_end = int(anchor_values[5])
                row_start = int(anchor_values[2])
                row_end = int(anchor_values[6])
                y_offset_end = int(anchor_values[7])

                # Calculate column span and width
                col_span = col_end - col_start
                width = col_span * AVG_COL_WIDTH_PT + (x_offset_end / 256.0) * AVG_COL_WIDTH_PT

                # Calculate row span and height
                row_span = row_end - row_start
                height = row_span * AVG_ROW_HEIGHT_PT + (y_offset_end / 256.0) * AVG_ROW_HEIGHT_PT

                return (width, height)
        except (ValueError, IndexError):
            pass

        return (None, None)
