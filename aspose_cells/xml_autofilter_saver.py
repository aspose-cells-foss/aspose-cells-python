"""
Aspose.Cells for Python - XML AutoFilter Saver Module

This module provides the AutoFilterXMLWriter class which handles saving
autofilter data to XML format according to ECMA-376 specification.

ECMA-376 Section 18.3.1.2 defines the autoFilter element structure.
"""


class AutoFilterXMLWriter:
    """
    Handles writing autofilter data to XML format for .xlsx files.

    The AutoFilterXMLWriter class is responsible for creating the XML
    representation of autofilter settings including filter columns,
    filter values, custom filters, color filters, dynamic filters,
    top10 filters, and sort state.

    Examples:
        >>> writer = AutoFilterXMLWriter()
        >>> xml = writer.format_auto_filter_xml(worksheet.auto_filter)
    """

    def __init__(self, escape_xml_func):
        """
        Initializes a new instance of the AutoFilterXMLWriter class.

        Args:
            escape_xml_func: A function to escape XML special characters.
        """
        self._escape_xml = escape_xml_func

    def format_auto_filter_xml(self, auto_filter):
        """
        Formats auto filter settings as XML according to ECMA-376 specification.

        ECMA-376 Section 18.3.1.2 defines the autoFilter element structure:
        - autoFilter: Main element with ref attribute (filter range)
        - filterColumn: Represents filter settings for a column (colId attribute)
        - filters: Contains filter values
        - filter: Individual filter value (val attribute)
        - customFilters: Contains custom filter criteria
        - customFilter: Custom filter criterion (operator and val attributes)
        - colorFilter: Color filter settings
        - dynamicFilter: Dynamic filter settings (type and val attributes)
        - top10: Top 10 filter settings (top, percent, and val attributes)
        - sortState: Sort state settings (columnOffset and order attributes)

        Args:
            auto_filter (AutoFilter): The AutoFilter object to format.

        Returns:
            str: XML representation of the auto filter settings.
        """
        xml = f'    <autoFilter ref="{auto_filter.range}">\n'

        # Write filter columns
        for col_id, filter_col in sorted(auto_filter.filter_columns.items()):
            xml += f'        <filterColumn colId="{col_id}"'

            # Add hiddenButton attribute if filter button is hidden
            if not filter_col.filter_button:
                xml += ' hiddenButton="1"'

            # Check if this column has any filters
            has_filters = (len(filter_col.filters) > 0 or
                          len(filter_col.custom_filters) > 0 or
                          filter_col.color_filter is not None or
                          filter_col.dynamic_filter is not None or
                          filter_col.top10_filter is not None)

            if not has_filters:
                xml += '/>\n'
                continue

            xml += '>\n'

            # Write filters (value filters)
            if len(filter_col.filters) > 0:
                xml += f'            <filters>\n'
                for filter_value in filter_col.filters:
                    escaped_value = self._escape_xml(str(filter_value))
                    xml += f'                <filter val="{escaped_value}"/>\n'
                xml += f'            </filters>\n'

            # Write custom filters
            if len(filter_col.custom_filters) > 0:
                xml += f'            <customFilters>\n'
                for operator, value in filter_col.custom_filters:
                    escaped_value = self._escape_xml(str(value))
                    xml += f'                <customFilter operator="{operator}" val="{escaped_value}"/>\n'
                xml += f'            </customFilters>\n'

            # Write color filter
            if filter_col.color_filter is not None:
                color_data = filter_col.color_filter
                cell_color_attr = ' cellColor="1"' if color_data['cell_color'] else ' cellColor="0"'
                xml += f'            <colorFilter rgb="{color_data["color"]}"{cell_color_attr}/>\n'

            # Write dynamic filter
            if filter_col.dynamic_filter is not None:
                dynamic_data = filter_col.dynamic_filter
                xml += f'            <dynamicFilter type="{dynamic_data["type"]}"'
                if dynamic_data['value'] is not None:
                    escaped_value = self._escape_xml(str(dynamic_data['value']))
                    xml += f' val="{escaped_value}"'
                xml += '/>\n'

            # Write top10 filter
            if filter_col.top10_filter is not None:
                top10_data = filter_col.top10_filter
                top_attr = ' top="1"' if top10_data['top'] else ' top="0"'
                percent_attr = ' percent="1"' if top10_data['percent'] else ' percent="0"'
                xml += f'            <top10{top_attr}{percent_attr} val="{top10_data["val"]}"/>\n'

            xml += f'        </filterColumn>\n'

        # Write sort state (ECMA-376 Section 18.3.1.92)
        if auto_filter.sort_state is not None and auto_filter.range is not None:
            sort_data = auto_filter.sort_state
            # Calculate sortState ref and sortCondition ref from autoFilter range
            sort_state_ref, sort_condition_ref = self._calculate_sort_refs(
                auto_filter.range, sort_data.get('column_index', 0)
            )
            xml += f'        <sortState ref="{sort_state_ref}">\n'
            # Build sortCondition with ref attribute
            descending_attr = ' descending="1"' if sort_data.get('descending', False) else ''
            xml += f'            <sortCondition ref="{sort_condition_ref}"{descending_attr}/>\n'
            xml += f'        </sortState>\n'

        xml += f'    </autoFilter>\n'

        return xml

    def _calculate_sort_refs(self, filter_range, column_index):
        """
        Calculates sortState ref and sortCondition ref from autoFilter range.

        According to ECMA-376:
        - sortState ref: The data range being sorted (excludes header row)
        - sortCondition ref: The specific column range being sorted

        Args:
            filter_range (str): The autoFilter range (e.g., "A1:D10")
            column_index (int): Zero-based column index within the filter range

        Returns:
            tuple: (sort_state_ref, sort_condition_ref)
        """
        from .cells import Cells

        # Parse the filter range
        if ':' in filter_range:
            start_ref, end_ref = filter_range.split(':')
        else:
            start_ref = end_ref = filter_range

        start_row, start_col = Cells.coordinate_from_string(start_ref)
        end_row, end_col = Cells.coordinate_from_string(end_ref)

        # sortState ref excludes header row (start from row after header)
        data_start_row = start_row + 1
        start_col_letter = Cells.column_letter_from_index(start_col)
        end_col_letter = Cells.column_letter_from_index(end_col)
        sort_state_ref = f"{start_col_letter}{data_start_row}:{end_col_letter}{end_row}"

        # sortCondition ref is for the specific column being sorted
        sort_col = start_col + column_index
        sort_col_letter = Cells.column_letter_from_index(sort_col)
        sort_condition_ref = f"{sort_col_letter}{data_start_row}:{sort_col_letter}{end_row}"

        return sort_state_ref, sort_condition_ref
