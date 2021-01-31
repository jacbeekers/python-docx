# encoding: utf-8

"""
Core properties part, corresponds to ``/docProps/core.xml`` part in package.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from datetime import datetime

from ..constants import CONTENT_TYPE as CT
from ..extendedprops import ExtendedProperties
from ...oxml.extendedprops import CT_ExtendedProperties
from ..packuri import PackURI
from ..part import XmlPart


class ExtendedPropertiesPart(XmlPart):
    """
    Corresponds to part named ``/docProps/app.xml``, containing the extended
    document properties for this document package.
    """

    @classmethod
    def default(cls, package):
        """
        Return a new |CorePropertiesPart| object initialized with default
        values for its base properties.
        """
        extended_properties_part = cls._new(package)
        extended_properties = extended_properties_part.extended_properties
        extended_properties.template = 'Normal.dotm'
        extended_properties.total_time = 0
        extended_properties.pages = 1
        extended_properties.words = 0
        extended_properties.characters = 0
        extended_properties.application = "python-docx"
        extended_properties.doc_security = 0
        extended_properties.lines = 1
        extended_properties.paragraphs = 1
        extended_properties.scale_crop = False
        # TODO: heading_pairs and titles_of_parts
        extended_properties.manager = "The Boss"
        extended_properties.company = ""
        extended_properties.links_up_to_date = False
        extended_properties.characters_with_spaces = 0
        extended_properties.shared_doc = False
        extended_properties.hyperlink_base = ""
        extended_properties.hyperlinks_changed = False
        extended_properties.app_version = "0.8.10"
        return extended_properties_part

    @property
    def extended_properties(self):
        """
        A |ExtendedProperties| object providing read/write access to the extended
        properties contained in this extended properties part.
        """
        return ExtendedProperties(self.element)

    @classmethod
    def _new(cls, package):
        partname = PackURI('/docProps/app.xml')
        content_type = CT.OFC_EXTENDED_PROPERTIES
        extendedProperties = CT_ExtendedProperties.new()
        return ExtendedPropertiesPart(
            partname, content_type, extendedProperties, package
        )
