"""
Aspose.Cells for Python - Hyperlink Module

This module provides hyperlink functionality for Excel workbooks,
compatible with the Excel object model.

ECMA-376 Compliance:
- Section 18.3.1.48: hyperlink element
- Section 18.3.1.49: hyperlinks collection
- Section 12.3.2.2: Hyperlink Relationships
"""


class Hyperlink:
    """
    Represents a hyperlink in a worksheet.

    A hyperlink can link to:
    - External web URLs (http, https, ftp)
    - Internal worksheet locations (Sheet2!A1)
    - Local or network files
    - Email addresses (mailto:)

    Examples:
        >>> # External URL
        >>> link = worksheet.hyperlinks.add("A1", "https://www.example.com")
        >>> link.text_to_display = "Visit Website"
        >>> link.screen_tip = "Click to visit our website"

        >>> # Internal link
        >>> link = worksheet.hyperlinks.add("B2", "", "Sheet2!A1")
        >>> link.text_to_display = "Go to Sheet2"
    """

    def __init__(self, range_address, address="", sub_address="", text_to_display="", screen_tip=""):
        """
        Initializes a new Hyperlink instance.

        Args:
            range_address (str): Cell reference (e.g., "A1" or "A1:B5").
            address (str, optional): External target URL. Defaults to "".
            sub_address (str, optional): Internal location (e.g., "Sheet2!A1"). Defaults to "".
            text_to_display (str, optional): Display text. Defaults to "".
            screen_tip (str, optional): Tooltip text. Defaults to "".
        """
        self._range = range_address
        self._address = address
        self._sub_address = sub_address
        self._text_to_display = text_to_display
        self._screen_tip = screen_tip
        self._relationship_id = None  # Set when saving to XML

    @property
    def range(self):
        """
        Gets the cell reference for this hyperlink.

        Returns:
            str: Cell reference (e.g., "A1" or "A1:B5").

        Examples:
            >>> link.range
            'A1'
        """
        return self._range

    @range.setter
    def range(self, value):
        """
        Sets the cell reference for this hyperlink.

        Args:
            value (str): Cell reference.
        """
        self._range = value

    @property
    def address(self):
        """
        Gets or sets the external target URL.

        For external hyperlinks (web URLs, files, email), this contains the target.
        For internal hyperlinks, use sub_address instead.

        Returns:
            str: External URL or empty string for internal links.

        Examples:
            >>> link.address = "https://www.example.com"
            >>> link.address
            'https://www.example.com'
        """
        return self._address

    @address.setter
    def address(self, value):
        """Sets the external target URL."""
        self._address = value if value else ""
        # Clear sub_address when setting external address
        if value:
            self._sub_address = ""

    @property
    def sub_address(self):
        """
        Gets or sets the internal location within the workbook.

        Format: "SheetName!CellReference" (e.g., "Sheet2!A1")
        Use single quotes for sheet names with spaces: "'My Sheet'!A1"

        Returns:
            str: Internal location or empty string for external links.

        Examples:
            >>> link.sub_address = "Sheet2!A1"
            >>> link.sub_address = "'Sales Data'!B5"
        """
        return self._sub_address

    @sub_address.setter
    def sub_address(self, value):
        """Sets the internal location."""
        self._sub_address = value if value else ""
        # Clear address when setting internal sub_address
        if value:
            self._address = ""

    @property
    def text_to_display(self):
        """
        Gets or sets the text displayed for the hyperlink.

        This is the text shown in the cell instead of the URL.
        If not set, the cell's value is displayed.

        Returns:
            str: Display text.

        Examples:
            >>> link.text_to_display = "Click Here"
        """
        return self._text_to_display

    @text_to_display.setter
    def text_to_display(self, value):
        """Sets the display text."""
        self._text_to_display = value if value else ""

    @property
    def screen_tip(self):
        """
        Gets or sets the tooltip text shown when hovering over the hyperlink.

        Returns:
            str: Tooltip text.

        Examples:
            >>> link.screen_tip = "Click to visit our website"
        """
        return self._screen_tip

    @screen_tip.setter
    def screen_tip(self, value):
        """Sets the tooltip text."""
        self._screen_tip = value if value else ""

    @property
    def type(self):
        """
        Gets the type of hyperlink.

        Returns:
            str: "External" for URLs/files/email, "Internal" for worksheet locations.

        Examples:
            >>> link.type
            'External'
        """
        if self._sub_address:
            return "Internal"
        elif self._address:
            return "External"
        else:
            return "None"

    def delete(self):
        """
        Marks this hyperlink for deletion.

        The hyperlink will be removed when the worksheet is saved.
        Note: This should be called from the Hyperlinks collection's delete method.
        """
        self._deleted = True

    def __repr__(self):
        """String representation of the hyperlink."""
        if self._sub_address:
            target = f"Internal: {self._sub_address}"
        elif self._address:
            target = f"External: {self._address}"
        else:
            target = "No target"
        return f"<Hyperlink range={self._range} {target}>"


class Hyperlinks:
    """
    Collection of hyperlinks in a worksheet.

    Provides methods to add, delete, and iterate over hyperlinks.

    Examples:
        >>> # Add external hyperlink
        >>> link = worksheet.hyperlinks.add("A1", "https://www.example.com")

        >>> # Add internal hyperlink
        >>> link = worksheet.hyperlinks.add("B2", "", "Sheet2!A1")

        >>> # Iterate over hyperlinks
        >>> for link in worksheet.hyperlinks:
        ...     print(link.range, link.address)

        >>> # Get hyperlink count
        >>> count = worksheet.hyperlinks.count
    """

    def __init__(self, worksheet):
        """
        Initializes a new Hyperlinks collection.

        Args:
            worksheet: The worksheet that owns this collection.
        """
        self._worksheet = worksheet
        self._hyperlinks = []

    def add(self, range_address, address="", sub_address="", text_to_display="", screen_tip=""):
        """
        Adds a new hyperlink to the collection.

        Args:
            range_address (str): Cell reference (e.g., "A1" or "A1:B5").
            address (str, optional): External URL for external links. Defaults to "".
            sub_address (str, optional): Internal location for internal links (e.g., "Sheet2!A1"). Defaults to "".
            text_to_display (str, optional): Display text. Defaults to "".
            screen_tip (str, optional): Tooltip text. Defaults to "".

        Returns:
            Hyperlink: The newly created hyperlink.

        Raises:
            ValueError: If both address and sub_address are provided, or if neither is provided.

        Examples:
            >>> # External web link
            >>> link = hyperlinks.add("A1", "https://www.example.com", text_to_display="Visit Website")

            >>> # Internal link
            >>> link = hyperlinks.add("B2", sub_address="Sheet2!A1", text_to_display="Go to Sheet2")

            >>> # Email link
            >>> link = hyperlinks.add("C3", "mailto:user@example.com", text_to_display="Email Us")

            >>> # File link
            >>> link = hyperlinks.add("D4", "file:///C:/Documents/report.pdf", text_to_display="Open Report")
        """
        # Validate parameters
        if address and sub_address:
            raise ValueError("Cannot specify both address and sub_address. Use one or the other.")
        if not address and not sub_address:
            raise ValueError("Must specify either address (external) or sub_address (internal).")

        # Create new hyperlink
        hyperlink = Hyperlink(
            range_address=range_address,
            address=address,
            sub_address=sub_address,
            text_to_display=text_to_display,
            screen_tip=screen_tip
        )

        self._hyperlinks.append(hyperlink)
        return hyperlink

    def delete(self, index=None, hyperlink=None):
        """
        Deletes a hyperlink from the collection.

        Args:
            index (int, optional): Zero-based index of the hyperlink to delete.
            hyperlink (Hyperlink, optional): The hyperlink object to delete.

        Raises:
            ValueError: If neither index nor hyperlink is provided.
            IndexError: If index is out of range.

        Examples:
            >>> # Delete by index
            >>> hyperlinks.delete(index=0)

            >>> # Delete by object
            >>> link = hyperlinks[0]
            >>> hyperlinks.delete(hyperlink=link)
        """
        if index is not None:
            if 0 <= index < len(self._hyperlinks):
                del self._hyperlinks[index]
            else:
                raise IndexError(f"Hyperlink index {index} out of range")
        elif hyperlink is not None:
            if hyperlink in self._hyperlinks:
                self._hyperlinks.remove(hyperlink)
            else:
                raise ValueError("Hyperlink not found in collection")
        else:
            raise ValueError("Must specify either index or hyperlink parameter")

    def clear(self):
        """
        Removes all hyperlinks from the collection.

        Examples:
            >>> worksheet.hyperlinks.clear()
        """
        self._hyperlinks.clear()

    @property
    def count(self):
        """
        Gets the number of hyperlinks in the collection.

        Returns:
            int: Number of hyperlinks.

        Examples:
            >>> count = worksheet.hyperlinks.count
            >>> print(f"There are {count} hyperlinks")
        """
        return len(self._hyperlinks)

    def __len__(self):
        """Returns the number of hyperlinks in the collection."""
        return len(self._hyperlinks)

    def __getitem__(self, index):
        """
        Gets a hyperlink by index.

        Args:
            index (int): Zero-based index.

        Returns:
            Hyperlink: The hyperlink at the specified index.

        Examples:
            >>> link = worksheet.hyperlinks[0]
        """
        return self._hyperlinks[index]

    def __iter__(self):
        """
        Returns an iterator over the hyperlinks.

        Examples:
            >>> for link in worksheet.hyperlinks:
            ...     print(link.range)
        """
        return iter(self._hyperlinks)

    def __repr__(self):
        """String representation of the collection."""
        return f"<Hyperlinks count={len(self._hyperlinks)}>"
