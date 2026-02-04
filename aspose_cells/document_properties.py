"""
Aspose.Cells for Python - Document Properties Module

This module provides classes for document-level properties according to ECMA-376 specification.
These properties are stored in docProps/core.xml and docProps/app.xml files.

ECMA-376 Part 2: Open Packaging Conventions - Core Properties
ECMA-376 Part 1: Extended Properties (docProps/app.xml)
"""

from datetime import datetime


class CoreProperties:
    """
    Represents core document properties stored in docProps/core.xml.

    Uses Dublin Core metadata elements and OPC core properties.

    ECMA-376 Part 2, Section 11 - Core Properties

    Examples:
        >>> wb.document_properties.core.title = "Sales Report"
        >>> wb.document_properties.core.creator = "John Doe"
        >>> wb.document_properties.core.subject = "Q4 2024 Sales"
    """

    def __init__(self):
        self._title = None
        self._subject = None
        self._creator = None
        self._keywords = None
        self._description = None
        self._last_modified_by = None
        self._revision = None
        self._created = None
        self._modified = None
        self._category = None
        self._content_status = None
        self._content_type = None
        self._identifier = None
        self._language = None
        self._version = None

    @property
    def title(self):
        """Document title (dc:title)."""
        return self._title

    @title.setter
    def title(self, value):
        self._title = value

    @property
    def subject(self):
        """Document subject (dc:subject)."""
        return self._subject

    @subject.setter
    def subject(self, value):
        self._subject = value

    @property
    def creator(self):
        """Document creator/author (dc:creator)."""
        return self._creator

    @creator.setter
    def creator(self, value):
        self._creator = value

    @property
    def keywords(self):
        """Keywords associated with the document (cp:keywords)."""
        return self._keywords

    @keywords.setter
    def keywords(self, value):
        self._keywords = value

    @property
    def description(self):
        """Document description/comments (dc:description)."""
        return self._description

    @description.setter
    def description(self, value):
        self._description = value

    @property
    def last_modified_by(self):
        """Name of person who last modified the document (cp:lastModifiedBy)."""
        return self._last_modified_by

    @last_modified_by.setter
    def last_modified_by(self, value):
        self._last_modified_by = value

    @property
    def revision(self):
        """Revision number (cp:revision)."""
        return self._revision

    @revision.setter
    def revision(self, value):
        self._revision = value

    @property
    def created(self):
        """Document creation date (dcterms:created)."""
        return self._created

    @created.setter
    def created(self, value):
        self._created = value

    @property
    def modified(self):
        """Document last modification date (dcterms:modified)."""
        return self._modified

    @modified.setter
    def modified(self, value):
        self._modified = value

    @property
    def category(self):
        """Document category (cp:category)."""
        return self._category

    @category.setter
    def category(self, value):
        self._category = value

    @property
    def content_status(self):
        """Content status such as Draft, Final (cp:contentStatus)."""
        return self._content_status

    @content_status.setter
    def content_status(self, value):
        self._content_status = value


class ExtendedProperties:
    """
    Represents extended/application properties stored in docProps/app.xml.

    ECMA-376 Part 1, Section 22.2 - Extended Properties

    Examples:
        >>> wb.document_properties.extended.application = "Microsoft Excel"
        >>> wb.document_properties.extended.company = "Acme Corp"
    """

    def __init__(self):
        self._application = "Microsoft Excel"
        self._app_version = None
        self._company = None
        self._manager = None
        self._doc_security = 0
        self._hyperlink_base = None
        self._scale_crop = False
        self._links_up_to_date = False
        self._shared_doc = False

    @property
    def application(self):
        """Name of the application that created the document."""
        return self._application

    @application.setter
    def application(self, value):
        self._application = value

    @property
    def app_version(self):
        """Version of the application that created the document."""
        return self._app_version

    @app_version.setter
    def app_version(self, value):
        self._app_version = value

    @property
    def company(self):
        """Company or organization name."""
        return self._company

    @company.setter
    def company(self, value):
        self._company = value

    @property
    def manager(self):
        """Manager associated with the document."""
        return self._manager

    @manager.setter
    def manager(self, value):
        self._manager = value

    @property
    def doc_security(self):
        """Document security level (0=none, 1=password protected, etc.)."""
        return self._doc_security

    @doc_security.setter
    def doc_security(self, value):
        self._doc_security = value

    @property
    def hyperlink_base(self):
        """Base URL for relative hyperlinks."""
        return self._hyperlink_base

    @hyperlink_base.setter
    def hyperlink_base(self, value):
        self._hyperlink_base = value

    @property
    def scale_crop(self):
        """Whether to scale or crop document thumbnail."""
        return self._scale_crop

    @scale_crop.setter
    def scale_crop(self, value):
        self._scale_crop = value

    @property
    def links_up_to_date(self):
        """Whether hyperlinks are up to date."""
        return self._links_up_to_date

    @links_up_to_date.setter
    def links_up_to_date(self, value):
        self._links_up_to_date = value

    @property
    def shared_doc(self):
        """Whether the document is shared."""
        return self._shared_doc

    @shared_doc.setter
    def shared_doc(self, value):
        self._shared_doc = value


class DocumentProperties:
    """
    Container for all document-level properties.

    This includes both core properties (docProps/core.xml) and
    extended properties (docProps/app.xml).

    Examples:
        >>> wb.document_properties.core.title = "Sales Report"
        >>> wb.document_properties.core.creator = "John Doe"
        >>> wb.document_properties.extended.company = "Acme Corp"
    """

    def __init__(self):
        self._core = CoreProperties()
        self._extended = ExtendedProperties()

    @property
    def core(self):
        """Gets core document properties (stored in docProps/core.xml)."""
        return self._core

    @property
    def extended(self):
        """Gets extended/application properties (stored in docProps/app.xml)."""
        return self._extended

    # Convenience properties that map to core properties
    @property
    def title(self):
        """Document title."""
        return self._core.title

    @title.setter
    def title(self, value):
        self._core.title = value

    @property
    def subject(self):
        """Document subject."""
        return self._core.subject

    @subject.setter
    def subject(self, value):
        self._core.subject = value

    @property
    def author(self):
        """Document author (alias for creator)."""
        return self._core.creator

    @author.setter
    def author(self, value):
        self._core.creator = value

    @property
    def creator(self):
        """Document creator."""
        return self._core.creator

    @creator.setter
    def creator(self, value):
        self._core.creator = value

    @property
    def keywords(self):
        """Document keywords."""
        return self._core.keywords

    @keywords.setter
    def keywords(self, value):
        self._core.keywords = value

    @property
    def comments(self):
        """Document comments (alias for description)."""
        return self._core.description

    @comments.setter
    def comments(self, value):
        self._core.description = value

    @property
    def category(self):
        """Document category."""
        return self._core.category

    @category.setter
    def category(self, value):
        self._core.category = value

    @property
    def company(self):
        """Company name."""
        return self._extended.company

    @company.setter
    def company(self, value):
        self._extended.company = value

    @property
    def manager(self):
        """Manager name."""
        return self._extended.manager

    @manager.setter
    def manager(self, value):
        self._extended.manager = value
