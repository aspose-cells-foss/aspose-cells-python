"""
Aspose.Cells for Python - XML Conditional Formatting Saver Module

This module provides the ConditionalFormatXMLWriter class which handles saving
conditional formatting data to XML format according to ECMA-376 specification.

ECMA-376 Section 18.3.1.18 defines the conditionalFormatting element structure.
"""


class ConditionalFormatXMLWriter:
    """
    Handles writing conditional formatting data to XML format for .xlsx files.

    The ConditionalFormatXMLWriter class is responsible for creating the XML
    representation of conditional formatting rules including cell value rules,
    text rules, date rules, formula rules, color scales, data bars, icon sets,
    and more.

    Examples:
        >>> writer = ConditionalFormatXMLWriter(escape_xml_func)
        >>> xml = writer.format_conditional_formatting_xml(worksheet.conditional_formats)
    """

    def __init__(self, escape_xml_func):
        """
        Initializes a new instance of the ConditionalFormatXMLWriter class.

        Args:
            escape_xml_func: A function to escape XML special characters.
        """
        self._escape_xml = escape_xml_func

    def format_conditional_formatting_xml(self, conditional_formats):
        """
        Formats conditional formatting as XML according to ECMA-376 specification.

        ECMA-376 Section 18.3.1.18 defines the conditionalFormatting element structure:
        - conditionalFormatting: Main element with sqref attribute (cell ranges)
        - cfRule: Individual conditional formatting rule

        Args:
            conditional_formats (ConditionalFormatCollection): The collection of conditional formats.

        Returns:
            str: XML representation of the conditional formatting.
        """
        xml = ''

        # Group conditional formats by range (sqref)
        formats_by_range = {}
        for cf in conditional_formats:
            if cf.range is None:
                continue
            if cf.range not in formats_by_range:
                formats_by_range[cf.range] = []
            formats_by_range[cf.range].append(cf)

        # Write conditionalFormatting elements
        for sqref, formats in formats_by_range.items():
            xml += f'    <conditionalFormatting sqref="{sqref}">\n'

            for cf in formats:
                xml += self._format_cf_rule_xml(cf)

            xml += '    </conditionalFormatting>\n'

        return xml

    def _format_cf_rule_xml(self, cf):
        """
        Formats a single conditional formatting rule as XML.

        Args:
            cf (ConditionalFormat): The conditional format object.

        Returns:
            str: XML representation of the cfRule element.
        """
        # Map user-friendly type names to ECMA-376 XML type names
        type_map = {
            'cellValue': 'cellIs',
            'text': 'containsText',
            'date': 'timePeriod',
            'formula': 'expression',
            'duplicateValues': 'duplicateValues',
            'uniqueValues': 'uniqueValues',
            'top10': 'top10',
            'bottom10': 'top10',  # ECMA-376: both use 'top10' type, distinguished by bottom attribute
            'aboveAverage': 'aboveAverage',
            'belowAverage': 'aboveAverage',  # ECMA-376: both use 'aboveAverage' type, distinguished by aboveAverage attribute
            'colorScale': 'colorScale',
            'dataBar': 'dataBar',
            'iconSet': 'iconSet',
            'containsText': 'containsText',
            'notContainsText': 'notContainsText',
            'beginsWith': 'beginsWith',
            'endsWith': 'endsWith',
            'containsBlanks': 'containsBlanks',
            'notContainsBlanks': 'notContainsBlanks',
            'containsErrors': 'containsErrors',
            'notContainsErrors': 'notContainsErrors',
            'timePeriod': 'timePeriod',
            'expression': 'expression',
            'cellIs': 'cellIs'
        }

        rule_type = type_map.get(cf._type, cf._type)
        if rule_type is None:
            return ''

        # Build cfRule attributes
        attrs = [f'type="{rule_type}"']

        # Add dxfId if formatting is applied (will be set during dxf registration)
        if hasattr(cf, '_dxf_id') and cf._dxf_id is not None:
            attrs.append(f'dxfId="{cf._dxf_id}"')

        attrs.append(f'priority="{cf.priority}"')

        if cf.stop_if_true:
            attrs.append('stopIfTrue="1"')

        # Add operator for cellIs type
        if rule_type == 'cellIs' and cf.operator:
            attrs.append(f'operator="{cf.operator}"')

        # Add text attribute for text-based rules
        if rule_type in ('containsText', 'notContainsText', 'beginsWith', 'endsWith'):
            if cf._formula1:
                escaped_text = self._escape_xml(str(cf._formula1))
                attrs.append(f'text="{escaped_text}"')

        # Add timePeriod attribute
        if rule_type == 'timePeriod' and cf._operator:
            attrs.append(f'timePeriod="{cf._operator}"')

        # Add top10 attributes
        if rule_type == 'top10':
            # Handle both 'top10' and 'bottom10' user types
            # Both map to ECMA-376 'top10' type, distinguished by bottom attribute
            if cf.top is not None:
                attrs.append('bottom="0"' if cf.top else 'bottom="1"')
            elif cf._type == 'bottom10':
                # For bottom10 type, explicitly set bottom="1"
                attrs.append('bottom="1"')
            if cf.percent is not None:
                attrs.append('percent="1"' if cf.percent else 'percent="0"')
            if cf.rank is not None:
                attrs.append(f'rank="{cf.rank}"')

        # Add aboveAverage attributes
        if rule_type == 'aboveAverage':
            # Handle both 'aboveAverage' and 'belowAverage' user types
            # Both map to ECMA-376 'aboveAverage' type, distinguished by aboveAverage attribute
            if cf.above is not None:
                attrs.append('aboveAverage="1"' if cf.above else 'aboveAverage="0"')
            elif cf._type == 'belowAverage':
                # For belowAverage type, explicitly set aboveAverage="0"
                attrs.append('aboveAverage="0"')
            if cf.std_dev is not None:
                attrs.append(f'stdDev="{cf.std_dev}"')

        xml = f'        <cfRule {" ".join(attrs)}'

        # Check if we need child elements
        has_children = False
        children_xml = ''

        # Add formula elements based on rule type
        if rule_type == 'expression':
            # Expression (formula) rules use the _formula property
            formula_value = cf._formula if cf._formula is not None else cf._formula1
            if formula_value is not None:
                formula_text = str(formula_value)
                # Remove leading '=' if present
                if formula_text.startswith('='):
                    formula_text = formula_text[1:]
                escaped_formula = self._escape_xml(formula_text)
                children_xml += f'            <formula>{escaped_formula}</formula>\n'
                has_children = True
        elif rule_type == 'cellIs':
            # Cell value rules use _formula1 and _formula2
            if cf._formula1 is not None:
                formula_text = str(cf._formula1)
                if formula_text.startswith('='):
                    formula_text = formula_text[1:]
                escaped_formula = self._escape_xml(formula_text)
                children_xml += f'            <formula>{escaped_formula}</formula>\n'
                has_children = True
            if cf._formula2 is not None:
                formula_text = str(cf._formula2)
                if formula_text.startswith('='):
                    formula_text = formula_text[1:]
                escaped_formula = self._escape_xml(formula_text)
                children_xml += f'            <formula>{escaped_formula}</formula>\n'
                has_children = True
        elif rule_type in ('containsText', 'notContainsText', 'beginsWith', 'endsWith'):
            # Text rules need a proper Excel formula
            # Get the first cell of the range to build the formula
            text_value = cf._formula1 if cf._formula1 is not None else ''
            first_cell = self._get_first_cell_from_range(cf.range)
            formula_text = self._build_text_rule_formula(rule_type, text_value, first_cell)
            if formula_text:
                escaped_formula = self._escape_xml(formula_text)
                children_xml += f'            <formula>{escaped_formula}</formula>\n'
                has_children = True

        # Add colorScale element
        if rule_type == 'colorScale':
            children_xml += self._format_color_scale_xml(cf)
            has_children = True

        # Add dataBar element
        if rule_type == 'dataBar':
            children_xml += self._format_data_bar_xml(cf)
            has_children = True

        # Add iconSet element
        if rule_type == 'iconSet':
            children_xml += self._format_icon_set_xml(cf)
            has_children = True

        if has_children:
            xml += '>\n'
            xml += children_xml
            xml += '        </cfRule>\n'
        else:
            xml += '/>\n'

        return xml

    def _format_color_scale_xml(self, cf):
        """Formats a colorScale element for conditional formatting."""
        xml = '            <colorScale>\n'

        # Determine if 2-color or 3-color scale
        is_3_color = cf.mid_color is not None or cf.color_scale_type == '3-color'

        # Add cfvo elements (conditional format value objects)
        xml += '                <cfvo type="min"/>\n'
        if is_3_color:
            xml += '                <cfvo type="percentile" val="50"/>\n'
        xml += '                <cfvo type="max"/>\n'

        # Add color elements
        min_color = cf.min_color or 'FFF8696B'  # Default red
        max_color = cf.max_color or 'FF63BE7B'  # Default green
        mid_color = cf.mid_color or 'FFFFEB84'  # Default yellow

        xml += f'                <color rgb="{min_color}"/>\n'
        if is_3_color:
            xml += f'                <color rgb="{mid_color}"/>\n'
        xml += f'                <color rgb="{max_color}"/>\n'

        xml += '            </colorScale>\n'
        return xml

    def _format_data_bar_xml(self, cf):
        """Formats a dataBar element for conditional formatting."""
        xml = '            <dataBar>\n'

        # Add cfvo elements
        xml += '                <cfvo type="min"/>\n'
        xml += '                <cfvo type="max"/>\n'

        # Add color element
        bar_color = cf.bar_color or 'FF638EC6'  # Default blue
        xml += f'                <color rgb="{bar_color}"/>\n'

        xml += '            </dataBar>\n'
        return xml

    def _format_icon_set_xml(self, cf):
        """Formats an iconSet element for conditional formatting."""
        icon_set_type = cf.icon_set_type or '3TrafficLights1'

        attrs = [f'iconSet="{icon_set_type}"']
        if cf.reverse_icons:
            attrs.append('reverse="1"')
        if cf.show_icon_only:
            attrs.append('showValue="0"')

        xml = f'            <iconSet {" ".join(attrs)}>\n'

        # Add cfvo elements based on icon set type
        # Determine number of icons
        num_icons = 3
        if icon_set_type.startswith('4'):
            num_icons = 4
        elif icon_set_type.startswith('5'):
            num_icons = 5

        # Add cfvo elements with percent thresholds
        for i in range(num_icons):
            percent_val = int(100 * i / num_icons)
            xml += f'                <cfvo type="percent" val="{percent_val}"/>\n'

        xml += '            </iconSet>\n'
        return xml

    def _get_first_cell_from_range(self, cell_range):
        """
        Gets the first cell reference from a range.

        Args:
            cell_range (str): Cell range (e.g., "A1:A10" or "A1")

        Returns:
            str: First cell reference (e.g., "A1")
        """
        if not cell_range:
            return 'A1'
        if ':' in cell_range:
            return cell_range.split(':')[0]
        return cell_range

    def _build_text_rule_formula(self, rule_type, text_value, first_cell):
        """
        Builds the Excel formula for text-based conditional formatting rules.

        According to ECMA-376, text rules require specific Excel formulas:
        - containsText: NOT(ISERROR(SEARCH("text",A1)))
        - notContainsText: ISERROR(SEARCH("text",A1))
        - beginsWith: LEFT(A1,LEN("text"))="text"
        - endsWith: RIGHT(A1,LEN("text"))="text"

        Args:
            rule_type (str): The rule type
            text_value (str): The text to search for
            first_cell (str): The first cell of the range

        Returns:
            str: The Excel formula
        """
        if not text_value:
            return ''

        # Escape double quotes in text value for Excel formula
        escaped_text = str(text_value).replace('"', '""')

        if rule_type == 'containsText':
            return f'NOT(ISERROR(SEARCH("{escaped_text}",{first_cell})))'
        elif rule_type == 'notContainsText':
            return f'ISERROR(SEARCH("{escaped_text}",{first_cell}))'
        elif rule_type == 'beginsWith':
            return f'LEFT({first_cell},LEN("{escaped_text}"))="{escaped_text}"'
        elif rule_type == 'endsWith':
            return f'RIGHT({first_cell},LEN("{escaped_text}"))="{escaped_text}"'

        return ''
