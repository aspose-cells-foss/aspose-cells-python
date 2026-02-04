"""
Aspose.Cells for Python - XML Conditional Formatting Loader Module

This module provides the ConditionalFormatXMLLoader class which handles loading
conditional formatting data from XML format according to ECMA-376 specification.

ECMA-376 Section 18.3.1.18 defines the conditionalFormatting element structure.
"""


class ConditionalFormatXMLLoader:
    """
    Handles loading conditional formatting data from XML format for .xlsx files.

    The ConditionalFormatXMLLoader class is responsible for parsing the XML
    representation of conditional formatting rules including cell value rules,
    text rules, date rules, formula rules, color scales, data bars, icon sets,
    and more.

    Examples:
        >>> loader = ConditionalFormatXMLLoader(namespaces, workbook)
        >>> loader.load_conditional_formatting(worksheet, worksheet_root)
    """

    def __init__(self, namespaces, workbook):
        """
        Initializes a new instance of the ConditionalFormatXMLLoader class.

        Args:
            namespaces: XML namespaces dictionary for parsing.
            workbook: The workbook instance for accessing dxf styles.
        """
        self.ns = namespaces
        self.workbook = workbook

    def load_conditional_formatting(self, worksheet, worksheet_root):
        """
        Loads conditional formatting from worksheet XML according to ECMA-376 specification.

        ECMA-376 Section 18.3.1.18 defines the conditionalFormatting element and its children:
        - conditionalFormatting: Main element with sqref attribute
        - cfRule: Individual conditional formatting rule

        Args:
            worksheet (Worksheet): The worksheet object to load data into.
            worksheet_root: The XML root element of the worksheet.
        """
        from .conditional_format import ConditionalFormat

        # Find all conditionalFormatting elements
        cf_elements = worksheet_root.findall('.//main:conditionalFormatting', namespaces=self.ns)

        for cf_elem in cf_elements:
            sqref = cf_elem.get('sqref')
            if not sqref:
                continue

            # Load each cfRule within this conditionalFormatting element
            for rule_elem in cf_elem.findall('main:cfRule', namespaces=self.ns):
                cf = ConditionalFormat()
                cf._range = sqref

                # Parse rule attributes
                rule_type = rule_elem.get('type')
                cf._type = rule_type

                # Parse priority
                priority = rule_elem.get('priority')
                if priority:
                    cf._priority = int(priority)

                # Parse stopIfTrue
                stop_if_true = rule_elem.get('stopIfTrue')
                cf._stop_if_true = stop_if_true == '1'

                # Parse dxfId for loading formatting later
                dxf_id = rule_elem.get('dxfId')
                if dxf_id:
                    cf._dxf_id = int(dxf_id)

                # Parse operator for cellIs type
                operator = rule_elem.get('operator')
                if operator:
                    cf._operator = operator

                # Parse text attribute for text-based rules
                text = rule_elem.get('text')
                if text:
                    cf._formula1 = text

                # Parse timePeriod attribute
                time_period = rule_elem.get('timePeriod')
                if time_period:
                    cf._operator = time_period

                # Parse top10 attributes
                if rule_type == 'top10':
                    bottom = rule_elem.get('bottom')
                    cf._top = bottom != '1'
                    percent = rule_elem.get('percent')
                    cf._percent = percent == '1'
                    rank = rule_elem.get('rank')
                    if rank:
                        cf._rank = int(rank)

                # Parse aboveAverage attributes
                if rule_type == 'aboveAverage':
                    above_average = rule_elem.get('aboveAverage', '1')
                    cf._above = above_average != '0'
                    std_dev = rule_elem.get('stdDev')
                    if std_dev:
                        cf._std_dev = int(std_dev)

                # Parse formula elements based on rule type
                formula_elems = rule_elem.findall('main:formula', namespaces=self.ns)
                if rule_type == 'expression':
                    # Expression rules store formula in _formula property
                    if len(formula_elems) > 0 and formula_elems[0].text:
                        cf._formula = formula_elems[0].text
                elif rule_type == 'cellIs':
                    # Cell value rules use _formula1 and _formula2
                    if len(formula_elems) > 0 and formula_elems[0].text:
                        cf._formula1 = formula_elems[0].text
                    if len(formula_elems) > 1 and formula_elems[1].text:
                        cf._formula2 = formula_elems[1].text
                elif rule_type in ('containsText', 'notContainsText', 'beginsWith', 'endsWith'):
                    # Text rules: the text attribute is already parsed above
                    # The formula element contains the Excel formula, not the text value
                    # We use the text attribute value which was set earlier
                    pass
                else:
                    # Default: store in _formula1
                    if len(formula_elems) > 0 and formula_elems[0].text:
                        cf._formula1 = formula_elems[0].text
                    if len(formula_elems) > 1 and formula_elems[1].text:
                        cf._formula2 = formula_elems[1].text

                # Parse colorScale element
                color_scale_elem = rule_elem.find('main:colorScale', namespaces=self.ns)
                if color_scale_elem is not None:
                    self._load_color_scale(cf, color_scale_elem)

                # Parse dataBar element
                data_bar_elem = rule_elem.find('main:dataBar', namespaces=self.ns)
                if data_bar_elem is not None:
                    self._load_data_bar(cf, data_bar_elem)

                # Parse iconSet element
                icon_set_elem = rule_elem.find('main:iconSet', namespaces=self.ns)
                if icon_set_elem is not None:
                    self._load_icon_set(cf, icon_set_elem)

                # Add to worksheet's conditional formats
                worksheet.conditional_formats._formats.append(cf)

        # Load dxf formatting and apply to conditional formats
        self._apply_dxf_to_conditional_formats(worksheet)

    def _load_color_scale(self, cf, color_scale_elem):
        """Loads colorScale element data into conditional format."""
        # Get cfvo elements to determine if 2-color or 3-color
        cfvo_elems = color_scale_elem.findall('main:cfvo', namespaces=self.ns)
        cf._color_scale_type = '3-color' if len(cfvo_elems) >= 3 else '2-color'

        # Get color elements
        color_elems = color_scale_elem.findall('main:color', namespaces=self.ns)
        if len(color_elems) >= 1:
            cf._min_color = color_elems[0].get('rgb')
        if len(color_elems) >= 3:
            cf._mid_color = color_elems[1].get('rgb')
            cf._max_color = color_elems[2].get('rgb')
        elif len(color_elems) >= 2:
            cf._max_color = color_elems[1].get('rgb')

    def _load_data_bar(self, cf, data_bar_elem):
        """Loads dataBar element data into conditional format."""
        # Get color element
        color_elem = data_bar_elem.find('main:color', namespaces=self.ns)
        if color_elem is not None:
            cf._bar_color = color_elem.get('rgb')

    def _load_icon_set(self, cf, icon_set_elem):
        """Loads iconSet element data into conditional format."""
        cf._icon_set_type = icon_set_elem.get('iconSet', '3TrafficLights1')
        cf._reverse_icons = icon_set_elem.get('reverse') == '1'
        cf._show_icon_only = icon_set_elem.get('showValue') == '0'

    def _apply_dxf_to_conditional_formats(self, worksheet):
        """Applies differential formatting (dxf) to conditional formats."""
        # dxf styles are stored in workbook._dxf_styles after loading styles.xml
        if not hasattr(self.workbook, '_dxf_styles') or not self.workbook._dxf_styles:
            return

        for cf in worksheet.conditional_formats:
            if hasattr(cf, '_dxf_id') and cf._dxf_id is not None:
                if cf._dxf_id < len(self.workbook._dxf_styles):
                    dxf_data = self.workbook._dxf_styles[cf._dxf_id]
                    self._apply_dxf_data_to_cf(cf, dxf_data)

    def _apply_dxf_data_to_cf(self, cf, dxf_data):
        """Applies dxf data to a conditional format."""
        if 'font' in dxf_data:
            font = dxf_data['font']
            cf._font.bold = font.get('bold', False)
            cf._font.italic = font.get('italic', False)
            cf._font.underline = font.get('underline', False)
            cf._font.strikethrough = font.get('strikethrough', False)
            if 'color' in font:
                cf._font.color = font['color']

        if 'fill' in dxf_data:
            fill = dxf_data['fill']
            cf._fill.pattern_type = fill.get('pattern_type', 'solid')
            cf._fill.foreground_color = fill.get('fg_color', 'FFFFFFFF')
            cf._fill.background_color = fill.get('bg_color', 'FFFFFFFF')
