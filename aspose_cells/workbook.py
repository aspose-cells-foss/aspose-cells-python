"""
Aspose.Cells for Python - Workbook Module

This module provides Workbook class which represents an Excel workbook.
The Workbook class provides methods to manage worksheets, save and load Excel files,
and access workbook-level properties.

Compatible with Aspose.Cells for .NET API structure.
"""

import os
import tempfile
from enum import Enum, auto
from .worksheet import Worksheet
from .cell import Cell
from .style import Style
from .xml_loader import XMLLoader
from .xml_saver import XMLSaver
from .workbook_properties import WorkbookProperties
from .document_properties import DocumentProperties
from .cfb_handler import is_encrypted_file
from .xlsx_encryptor import XLSXEncryptor, XLSXDecryptor
from .csv_handler import CSVHandler, CSVLoadOptions, CSVSaveOptions
from .markdown_handler import MarkdownHandler, MarkdownSaveOptions
from .json_handler import JsonHandler, JsonSaveOptions


class SaveFormat(Enum):
    """
    Specifies the format for saving a workbook.

    Compatible with Aspose.Cells for .NET SaveFormat enumeration.

    Examples:
        >>> wb.save('output.xlsx', SaveFormat.XLSX)
        >>> wb.save('output.csv', SaveFormat.CSV)
        >>> wb.save('output.md', SaveFormat.MARKDOWN)
    """
    AUTO = auto()       # Auto-detect format from file extension
    XLSX = auto()       # Excel 2007+ format (.xlsx)
    CSV = auto()        # Comma-separated values (.csv)
    TSV = auto()        # Tab-separated values (.tsv)
    MARKDOWN = auto()   # Markdown format (.md)
    JSON = auto()       # JSON format (.json)

    @classmethod
    def from_extension(cls, file_path):
        """
        Determines the save format from file extension.

        Args:
            file_path (str): Path to the file.

        Returns:
            SaveFormat: The appropriate save format.

        Raises:
            ValueError: If file extension is not supported.
        """
        ext = os.path.splitext(file_path)[1].lower()
        format_map = {
            '.xlsx': cls.XLSX,
            '.xlsm': cls.XLSX,
            '.csv': cls.CSV,
            '.tsv': cls.TSV,
            '.md': cls.MARKDOWN,
            '.markdown': cls.MARKDOWN,
            '.json': cls.JSON,
        }
        if ext in format_map:
            return format_map[ext]
        raise ValueError(f"Unsupported file extension: {ext}")


class Workbook:
    """
    Represents an Excel workbook.
    
    The Workbook class provides methods and properties for working with Excel workbooks,
    including worksheet management, file I/O operations, and workbook-level settings.
    
    Examples:
        >>> wb = Workbook()
        >>> ws = wb.worksheets[0]
        >>> ws.cells['A1'].value = "Hello"
        >>> wb.save('output.xlsx')
        
        >>> wb = Workbook('input.xlsx')
        >>> print(wb.worksheets[0].cells['A1'].value)
    """
    
    def __init__(self, file_path=None, password=None):
        """
        Initializes a new instance of Workbook class.

        Args:
            file_path (str, optional): Path to an existing Excel file to load.
                                        If None, creates a new empty workbook.
            password (str, optional): Password for encrypted files.

        Examples:
            >>> wb = Workbook()  # Create new workbook
            >>> wb = Workbook('existing.xlsx')  # Load existing workbook
            >>> wb = Workbook('encrypted.xlsx', password='SecurePass123')  # Load encrypted workbook
        """
        self._worksheets = []
        self._styles = []
        self._shared_strings = []
        self._file_path = file_path

        # Workbook properties
        self._properties = WorkbookProperties()

        # Document properties
        self._document_properties = None

        # Style management (initialized by XMLSaver when saving)
        self._font_styles = {}  # Map font tuples to style indices
        self._fill_styles = {}  # Map fill tuples to style indices
        self._border_styles = {}  # Map border tuples to style indices
        self._alignment_styles = {}  # Map alignment tuples to style indices
        self._protection_styles = {}  # Map protection tuples to style indices
        self._cell_styles = {}  # Map cell style tuples to xf indices
        self._num_formats = {}  # Map number format strings to format IDs

        # Initialize with default style
        default_style = Style()
        self._styles.append(default_style)

        if file_path and os.path.exists(file_path):
            self._load(file_path, password)
        else:
            # Create default worksheet
            self._worksheets.append(Worksheet("Sheet1"))
    
    # Properties
    
    @property
    def worksheets(self):
        """
        Gets collection of worksheets in the workbook.
        
        Returns:
            list: List of Worksheet objects.
            
        Examples:
            >>> for ws in wb.worksheets:
            ...     print(ws.name)
        """
        return self._worksheets
    
    @property
    def file_path(self):
        """
        Gets the file path of the workbook.
        
        Returns:
            str: The file path, or None if workbook was not loaded from a file.
        """
        return self._file_path
    
    @property
    def properties(self):
        """
        Gets workbook properties.
        
        Returns:
            WorkbookProperties: The workbook properties object containing
            file version, protection, view, calculation, and defined names.
            
        Examples:
            >>> wb.properties.view.active_tab = 0
            >>> wb.properties.protection.lock_structure = True
            >>> wb.properties.calculation.calc_mode = "auto"
        """
        return self._properties
    
    @property
    def document_properties(self):
        """
        Gets document properties of the workbook.

        Returns:
            DocumentProperties: The document properties object containing core and extended properties.
        """
        if self._document_properties is None:
            self._document_properties = DocumentProperties()
        return self._document_properties
    
    # Worksheet management methods
    
    def add_worksheet(self, name=None):
        """
        Adds a new worksheet to the workbook.
        
        Args:
            name (str, optional): Name for the new worksheet. If None, a default name is generated.
            
        Returns:
            Worksheet: The newly created Worksheet object.
            
        Examples:
            >>> ws = wb.add_worksheet("NewSheet")
            >>> ws = wb.add_worksheet()  # Auto-generated name
        """
        if name is None:
            # Generate default name
            existing_names = [ws.name for ws in self._worksheets]
            i = 1
            while f"Sheet{i}" in existing_names:
                i += 1
            name = f"Sheet{i}"
        
        worksheet = Worksheet(name)
        self._worksheets.append(worksheet)
        return worksheet
    
    def get_worksheet(self, index_or_name):
        """
        Gets a worksheet by index or name.
        
        Args:
            index_or_name: Either an integer index (0-based) or string name of the worksheet.
            
        Returns:
            Worksheet: The Worksheet object at the specified index or with the specified name.
            
        Raises:
            IndexError: If index is out of range.
            ValueError: If no worksheet with the specified name exists.
            
        Examples:
            >>> ws = wb.get_worksheet(0)  # Get first worksheet by index
            >>> ws = wb.get_worksheet("Sheet2")  # Get worksheet by name
        """
        if isinstance(index_or_name, int):
            if 0 <= index_or_name < len(self._worksheets):
                return self._worksheets[index_or_name]
            raise IndexError(f"Worksheet index {index_or_name} out of range")
        elif isinstance(index_or_name, str):
            for ws in self._worksheets:
                if ws.name == index_or_name:
                    return ws
            raise ValueError(f"Worksheet '{index_or_name}' not found")
        else:
            raise TypeError("index_or_name must be int or str")
    
    def remove_worksheet(self, index_or_name):
        """
        Removes a worksheet from the workbook.
        
        Args:
            index_or_name: Either an integer index (0-based) or string name of the worksheet.
            
        Examples:
            >>> wb.remove_worksheet(0)  # Remove first worksheet
            >>> wb.remove_worksheet("Sheet2")  # Remove worksheet by name
        """
        if isinstance(index_or_name, int):
            if 0 <= index_or_name < len(self._worksheets):
                self._worksheets.pop(index_or_name)
            else:
                raise IndexError(f"Worksheet index {index_or_name} out of range")
        elif isinstance(index_or_name, str):
            for i, ws in enumerate(self._worksheets):
                if ws.name == index_or_name:
                    self._worksheets.pop(i)
                    return
            raise ValueError(f"Worksheet '{index_or_name}' not found")
        else:
            raise TypeError("index_or_name must be int or str")
    
    # File I/O methods

    def save(self, file_path, save_format=None, options=None, password=None, encryption_params=None):
        """
        Saves the workbook to a file.

        The file format is determined by the file extension or the explicit save_format parameter.
        Supported formats: XLSX, CSV, TSV, Markdown.

        Args:
            file_path (str): Path where the file should be saved.
            save_format (SaveFormat, optional): Explicit format specification.
                If None, format is auto-detected from file extension.
            options: Format-specific save options (CSVSaveOptions, MarkdownSaveOptions, etc.).
            password (str, optional): Password to encrypt the file (XLSX only).
            encryption_params (EncryptionParameters, optional): Encryption parameters (XLSX only).

        Examples:
            Auto-detect format from extension::

                from aspose_cells import Workbook, SaveFormat

                wb = Workbook()
                # ... add data ...

                wb.save('output.xlsx')  # Excel format
                wb.save('output.csv')   # CSV format
                wb.save('output.tsv')   # TSV format
                wb.save('output.md')    # Markdown format

            Explicit format specification::

                wb.save('data.txt', SaveFormat.CSV)       # Save .txt as CSV
                wb.save('report.txt', SaveFormat.MARKDOWN)  # Save .txt as Markdown

            With encryption (XLSX only)::

                wb.save('output.xlsx', password='SecurePass123')

            With format-specific options::

                from aspose_cells import MarkdownSaveOptions
                options = MarkdownSaveOptions()
                options.include_worksheet_name = False
                wb.save('report.md', options=options)

        Raises:
            ValueError: If file extension is not supported and no save_format specified.
        """
        # Determine save format
        if save_format is None or save_format == SaveFormat.AUTO:
            save_format = SaveFormat.from_extension(file_path)

        # Dispatch to appropriate save method
        if save_format == SaveFormat.XLSX:
            self._save_xlsx(file_path, password, encryption_params)
        elif save_format == SaveFormat.CSV:
            self.save_as_csv(file_path, options)
        elif save_format == SaveFormat.TSV:
            # TSV is CSV with tab delimiter
            if options is None:
                options = CSVSaveOptions()
            options.delimiter = '\t'
            self.save_as_csv(file_path, options)
        elif save_format == SaveFormat.MARKDOWN:
            self.save_as_markdown(file_path, options)
        elif save_format == SaveFormat.JSON:
            self.save_as_json(file_path, options)
        else:
            raise ValueError(f"Unsupported save format: {save_format}")

    def _save_xlsx(self, file_path, password=None, encryption_params=None):
        """
        Saves the workbook to an Excel file (.xlsx format).

        Args:
            file_path (str): Path where the Excel file should be saved.
            password (str, optional): Password to encrypt the file.
            encryption_params (EncryptionParameters, optional): Encryption parameters.
        """
        saver = XMLSaver(self)
        # Register default styles before saving
        saver.register_default_styles()

        if password:
            # Save to temporary file first
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp_path = tmp.name

            try:
                # Save unencrypted to temp file
                saver.save(tmp_path)

                # Encrypt temp file to final destination
                encryptor = XLSXEncryptor(encryption_params)
                encryptor.encrypt_file(tmp_path, file_path, password)
            finally:
                # Clean up temp file
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
        else:
            # Save directly without encryption
            saver.save(file_path)
    
    def _load(self, file_path, password=None):
        """
        Loads a workbook from an Excel file.

        Args:
            file_path (str): Path to the Excel file to load.
            password (str, optional): Password for encrypted files.

        Raises:
            ValueError: If file is encrypted but no password provided, or password is incorrect.
        """
        import zipfile

        # Check if file is encrypted
        if is_encrypted_file(file_path):
            if not password:
                raise ValueError("File is encrypted. Please provide a password.")

            # Decrypt to temporary file
            decryptor = XLSXDecryptor()
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp_path = tmp.name

            try:
                decryptor.decrypt_file(file_path, tmp_path, password)

                # Load decrypted file
                with zipfile.ZipFile(tmp_path, 'r') as zipf:
                    loader = XMLLoader(self)
                    loader.load_workbook(zipf)
            finally:
                # Clean up temp file
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
        else:
            # Load unencrypted file directly
            with zipfile.ZipFile(file_path, 'r') as zipf:
                loader = XMLLoader(self)
                loader.load_workbook(zipf)
    
    # CSV I/O methods

    def save_as_csv(self, file_path, options=None):
        """
        Saves the workbook to a CSV file.

        Args:
            file_path (str): Path where the CSV file should be saved.
            options (CSVSaveOptions, optional): Export options. Uses defaults if None.

        Examples:
            >>> wb.save_as_csv('output.csv')
            >>> options = CSVSaveOptions()
            >>> options.delimiter = ';'
            >>> wb.save_as_csv('output.csv', options)
        """
        CSVHandler.save_csv(self, file_path, options)

    def load_csv(self, file_path, options=None):
        """
        Loads data from a CSV file into the workbook.

        The CSV data is loaded into the first worksheet, replacing any existing data.

        Args:
            file_path (str): Path to the CSV file to load.
            options (CSVLoadOptions, optional): Import options. Uses defaults if None.

        Examples:
            >>> wb = Workbook()
            >>> wb.load_csv('data.csv')
            >>> options = CSVLoadOptions()
            >>> options.delimiter = ';'
            >>> wb.load_csv('data.csv', options)
        """
        CSVHandler.load_csv(self, file_path, options)

    # Markdown export methods

    def save_as_markdown(self, file_path, options=None):
        """
        Saves the workbook to a Markdown file.

        Args:
            file_path (str): Path where the Markdown file should be saved.
            options (MarkdownSaveOptions, optional): Export options. Uses defaults if None.

        Examples:
            >>> wb.save_as_markdown('output.md')
            >>> options = MarkdownSaveOptions()
            >>> options.default_alignment = 'center'
            >>> wb.save_as_markdown('output.md', options)
        """
        MarkdownHandler.save_markdown(self, file_path, options)

    # JSON export methods

    def save_as_json(self, file_path, options=None):
        """
        Saves the workbook to a JSON file.

        Args:
            file_path (str): Path where the JSON file should be saved.
            options (JsonSaveOptions, optional): Export options. Uses defaults if None.

        Examples:
            >>> wb.save_as_json('output.json')
            >>> options = JsonSaveOptions()
            >>> options.worksheet_index = 0
            >>> wb.save_as_json('sheet1.json', options)
        """
        JsonHandler.save_json(self, file_path, options)

    # String representation

    def __repr__(self):
        """
        Returns a string representation of the workbook.

        Returns:
            str: String representation showing the number of worksheets.
        """
        return f"Workbook(worksheets={len(self._worksheets)})"
