# encoding: utf-8

"""
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)
from lxml.etree import tostring


class ExtendedProperties(object):
    """
    Corresponds to part named ``/docProps/app.xml``, containing the extended
    document properties for this document package.
    """
    def __init__(self, element):
        self._element = element
        self.extended_properties = {}
        for s in self._element:
            self.extended_properties.update({s.tag.split('}')[1]: s.text})
        # print("extended properties list:", self.extended_properties)
        self.template = self.extended_properties["Template"] if "Template" in self.extended_properties else ""
        self.total_time = self.extended_properties["TotalTime"] if "TotalTime" in self.extended_properties else ""
        self.pages = self.extended_properties["Pages"] if "Pages" in self.extended_properties else ""
        self.words = self.extended_properties["Words"] if "Words" in self.extended_properties else ""
        self.characters = self.extended_properties["Characters"] if "Characters" in self.extended_properties else ""
        self.application = self.extended_properties["Application"] if "Application" in self.extended_properties else ""
        self.doc_security = self.extended_properties["DocSecurity"] if "DocSecurity" in self.extended_properties else ""
        self.lines = self.extended_properties["Lines"] if "Lines" in self.extended_properties else ""
        self.paragraphs = self.extended_properties["Paragraphs"] if "Paragraphs" in self.extended_properties else ""
        self.scale_crop = self.extended_properties["ScaleCrop"] if "ScaleCrop" in self.extended_properties else ""
        self.total_time = self.extended_properties["TotalTime"] if "TotalTime" in self.extended_properties else ""
        # TODO: HeadingPairs and TitlesOfParts
        self.manager = self.extended_properties["Manager"] if "Manager" in self.extended_properties else ""
        self.company = self.extended_properties["Company"] if "Company" in self.extended_properties else ""
        self.links_up_to_date = self.extended_properties["LinksUpToDate"] if "LinksUpToDate" in self.extended_properties else ""
        self.characters_with_spaces = self.extended_properties["CharactersWithSpaces"] if "CharactersWithSpaces" in self.extended_properties else ""
        self.shared_doc = self.extended_properties["SharedDoc"] if "SharedDoc" in self.extended_properties else ""
        self.hyperlink_base = self.extended_properties["HyperlinkBase"] if "HyperlinkBase" in self.extended_properties else ""
        self.hyperlinks_changed = self.extended_properties["HyperlinksChanged"] if "HyperlinksChanged" in self.extended_properties else ""
        self.app_version = self.extended_properties["AppVersion"] if "AppVersion" in self.extended_properties else ""


