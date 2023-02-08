import xml.etree.ElementTree as ET
from typing import Any, List, Optional
from xml.etree.ElementTree import Element
from zipfile import ZipFile

from pydantic import Field, dataclasses

from exmlrd import log
from exmlrd.tools import del_namespace, set_classattr

logger = log.get_logger(__name__)


@dataclasses.dataclass
class RichText:
    sz: str = ""  # font size
    color: str = ""
    rFont: str = ""
    family: str = ""
    scheme: str = ""
    charset: str = ""
    text: str = ""


@dataclasses.dataclass
class SiTag:
    rpr: list[RichText] = Field(default_factory=list)

    @property
    def text(self) -> str:
        if len(self.rpr) < 1:
            return ""
        text = [r.text for r in self.rpr]
        return "".join(text)


class SharedStyle:

    sharedstyle_xml = "xl/sharedStrings.xml"

    def __init__(self, archive: ZipFile):
        self.archive = archive

        try:
            self.root_tree = ET.parse(self.archive.open(self.sharedstyle_xml)).getroot()
        except KeyError:
            self.root_tree = None
        self.si = self.__get_shareitem()

    def __get_shareitem(self) -> list[Optional[SiTag]]:
        if self.root_tree is None:
            si = SiTag()
            return []
        shareditems: list[Any] = []
        shared_appned = shareditems.append
        for _e_root in self.root_tree:
            si = SiTag()
            __text: str = ""
            for _e_parents in _e_root:
                attribute = {}
                if del_namespace(_e_parents.tag) == "r":
                    for _e_child in _e_parents:
                        tag = del_namespace(_e_child.tag)
                        if tag == "t":
                            attribute["text"] = (
                                _e_child.text if isinstance(_e_child.text, str) else ""
                            )
                        elif tag == "rPr":
                            for _e in _e_child:
                                _tag = del_namespace(_e.tag)
                                if _e.attrib.get("val"):
                                    attribute[_tag] = _e.attrib["val"]
                                if _e.attrib.get("theme"):
                                    attribute[_tag] = _e.attrib["theme"]
                        __rich = RichText(**attribute)
                    si.rpr.append(__rich)
                elif del_namespace(_e_parents.tag) == "t":
                    __text = _e_parents.text if isinstance(_e_parents.text, str) else ""
                    __rich = RichText(text=__text)
                    si.rpr.append(__rich)
            shared_appned(si)
        return shareditems

    def get_shareitem(self, index: int) -> SiTag:
        try:
            si = self.si[index]
        except IndexError as e:
            logger.debug(e)
            si = SiTag()

        if si is None:
            si = SiTag()

        return si
