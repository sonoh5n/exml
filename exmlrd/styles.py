import re
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple
from xml.etree.ElementTree import Element
from zipfile import ZipFile

from pydantic import Field, dataclasses

from exmlrd import log
from exmlrd.tags import StylesTag
from exmlrd.tools import del_namespace, set_classattr

logger = log.get_logger(__name__)


@dataclasses.dataclass
class Font:
    sz: str = ""
    name: str = ""
    family: str = ""
    charset: str = ""
    scheme: str = ""
    color: str = ""
    b: bool = False
    i: bool = False
    u: bool = False
    strike: bool = False
    outline: bool = False
    shadow: bool = False
    condense: str = ""
    extend: str = ""
    vertAlign: str = ""


@dataclasses.dataclass
class FgColor:
    theme: Optional[str] = ""
    rgb: Optional[str] = ""
    tint: Optional[str] = ""


@dataclasses.dataclass
class BgColor:
    indexed: str = ""


@dataclasses.dataclass
class Fills:
    patternFill: str = ""
    fgColor: Optional[FgColor] = None
    bgColor: Optional[BgColor] = None


@dataclasses.dataclass
class Border:
    left: bool = False
    right: bool = False
    top: bool = False
    bottom: bool = False
    diagonal: bool = False


@dataclasses.dataclass
class NumFmt:
    id: str = ""
    formatCode: str = ""

    @property
    def unit(self) -> str:
        _, u, _ = self.parse_formatcode(self.formatCode)
        if u is None:
            u = ""
        return u

    @property
    def number(self) -> str:
        n, _, _ = self.parse_formatcode(self.formatCode)
        if n is None:
            n = ""
        return n

    @property
    def general(self) -> str:
        _, _, g = self.parse_formatcode(self.formatCode)
        if g is None:
            g = ""
        return g

    def parse_formatcode(
        self, input_text: str
    ) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """This function splits a given string into 3 parts: number, unit, and the remaining text.

            Args:
                input_text (str): The string to be split.

            Returns:
                Tuple[Optional[str], Optional[str], Optional[str]]: A tuple of 3 elements containing the number, unit, and the remaining text.
        The number and unit may be `None` if they are not present in the input string.

            Example:
            >>> split_number_unit("0.000Ω")
            ["0.000", "Ω", None]
            >>> split_number_unit("= General cm2")
            [None, "cm2", "= General"]
            >>> split_number_unit("0.0 ")
            ["0.0", None, None]
        """
        if input_text.strip() == "":
            return None, None, None
        input_text = input_text[::-1]
        match = re.match(r"^([\d.E+-]+)([^\d]+)?(.*)", input_text)
        if match:
            num = match.group(1)[::-1]
            unit = match.group(2)[::-1] if match.group(2) else None
            remaining = match.group(3)[::-1]
            return num, unit, remaining
        return None, None, input_text[::-1]


@dataclasses.dataclass
class XFS:
    numFmtId: str = ""
    fontId: str = ""
    fillId: str = ""
    borderId: str = ""
    xfId: str = ""
    applyFont: str = ""
    applyAlignment: str = ""
    applyBorder: str = ""
    applyFill: str = ""
    applyNumberFormat: str = ""
    applyProtection: str = ""


@dataclasses.dataclass
class Format:
    numFmt: NumFmt = Field(default_factory=NumFmt)
    font: Font = Field(default_factory=Font)
    fill: Fills = Field(default_factory=Fills)
    border: Border = Field(default_factory=Border)


class Styels:
    """
    Class to retrieve cell styles

    Get cell formatting (font, font size, fill, border, etc.) from sytle.xml

    Attributes:
        root_tree (Element | Any): The root element in the xml structure of sytle.xml
        cellxfs (list[XFS]): Another class attribute with a description.
        fontid (list[Font]): All element information in font tag in sytle.xml
        fillid (list[Fills]): All element information in fill tag in sytle.xml
        borders (list[Border]): All element information in borders tag in sytle.xml
        numfmt (dict[str, NumFmt]): All element information in numfmt tag in sytle.xml
    """

    styles_xml = "xl/styles.xml"

    def __init__(self, archive: ZipFile):
        self.root_tree = ET.parse(archive.open(self.styles_xml)).getroot()
        self.cellxfs = self.__get_cellXfs()
        self.fontid = self.__get_fontid()
        self.fillid = self.__get_fillid()
        self.borders = self.__get_borders()
        self.numfmt = self.__get_numfmt()

    def __get_elem_index(self, elem: Element, word: str) -> Optional[int]:
        if not isinstance(elem, Element):
            return None
        else:
            for index, __e in enumerate(elem):
                if del_namespace(__e.tag) == word:
                    return index
        return None

    def __get_fontid(self) -> list[Font]:
        fontlists: list[Font] = []
        __eidx = self.__get_elem_index(self.root_tree, StylesTag.FONTS.value)
        if __eidx is None:
            return fontlists

        font_appned = fontlists.append
        for _e_parents in self.root_tree[__eidx]:
            font = Font()
            for _e_child_avalue in _e_parents:
                tag = del_namespace(_e_child_avalue.tag)
                if _e_child_avalue.attrib.get("val"):
                    set_classattr(font, key=tag, value=_e_child_avalue.attrib["val"])
                elif _e_child_avalue.attrib.get("theme"):
                    set_classattr(font, key=tag, value=_e_child_avalue.attrib["theme"])
                elif _e_child_avalue.attrib.get("rgb"):
                    set_classattr(font, key=tag, value=_e_child_avalue.attrib["rgb"])

                if tag in ["b", "i", "u", "strike", "outline", "shadow"]:
                    set_classattr(font, key=tag, value=True)
            font_appned(font)
        return fontlists

    def get_fontid(self, index: int) -> Font:
        try:
            item = self.fontid[index]
        except IndexError as e:
            logger.debug(e)
            item = Font()
        return item

    def __get_cellXfs(self) -> list[XFS]:
        cellXfslists: list[XFS] = []
        __eidx = self.__get_elem_index(self.root_tree, StylesTag.CELLXFS.value)
        if __eidx is None:
            return cellXfslists

        xfs_appned = cellXfslists.append
        for _e_parents in self.root_tree[__eidx]:
            if del_namespace(_e_parents.tag):
                xfs_appned(XFS(**_e_parents.attrib))
        return cellXfslists

    def get_cellXfs(self, index: int) -> XFS:
        try:
            cellxfs = self.cellxfs[index]
        except IndexError as e:
            logger.debug(e)
            cellxfs = XFS()
        return cellxfs

    def __get_fillid(self) -> list[Fills]:
        fillslists: list[Fills] = []
        __eidx = self.__get_elem_index(self.root_tree, StylesTag.FILLS.value)
        if __eidx is None:
            return fillslists

        fill_appned = fillslists.append
        for _e_parents in self.root_tree[__eidx]:
            fills = Fills()
            for _e_child in _e_parents:
                tag = del_namespace(_e_child.tag)
                if tag == "patternFill":
                    fills.patternFill = (
                        _e_child.attrib["patternType"]
                        if "patternType" in _e_child.attrib
                        else ""
                    )
                    for _e in _e_child:
                        tag = del_namespace(_e.tag)
                        if tag == "fgColor":
                            fgcolor = FgColor()
                            set_classattr(fgcolor, container=_e.attrib)
                            fills.fgColor = fgcolor
                        elif tag == "bgColor":
                            bgcolor = BgColor()
                            set_classattr(bgcolor, container=_e.attrib)
                            fills.bgColor = bgcolor
            fill_appned(fills)
        return fillslists

    def get_fillid(self, index: int) -> Fills:
        try:
            fills = self.fillid[index]
        except IndexError as e:
            logger.debug(e)
            fills = Fills()
        return fills

    def __get_borders(self) -> list[Border]:
        borderlists: list[Border] = []
        __eidx = self.__get_elem_index(self.root_tree, StylesTag.BORDERS.value)
        if __eidx is None:
            return borderlists

        borders_appned = borderlists.append
        for _e_parents in self.root_tree[__eidx]:
            borders = Border()
            for _e_child in _e_parents:
                tag = del_namespace(_e_child.tag)
                if tag == "left" and _e_child.get("style"):
                    borders.left = True
                elif tag == "right" and _e_child.get("style"):
                    borders.right = True
                elif tag == "top" and _e_child.get("style"):
                    borders.top = True
                elif tag == "bottom" and _e_child.get("style"):
                    borders.bottom = True
                elif tag == "diagonal" and _e_child.get("style"):
                    borders.diagonal = True
            borders_appned(borders)
        return borderlists

    def get_borders(self, index: int) -> Border:
        try:
            borders = self.borders[index]
        except IndexError as e:
            logger.debug(e)
            borders = Border()
        return borders

    def __get_numfmt(self) -> dict[str, NumFmt]:
        numfmts: dict[str, NumFmt] = {}
        __eidx = self.__get_elem_index(self.root_tree, StylesTag.NUMFMTS.value)
        if __eidx is None:
            return numfmts

        for _e_parents in self.root_tree[__eidx]:
            fid = _e_parents.attrib["numFmtId"]
            if fid:
                numfmts[fid] = NumFmt(**_e_parents.attrib)
        return numfmts

    def get_numfmt(self, index: int) -> NumFmt:
        try:
            numfmt = self.numfmt.get(str(index))
        except KeyError as e:
            logger.debug(e)
            numfmt = NumFmt()

        if numfmt is None:
            numfmt = NumFmt()

        return numfmt

    def get_format(self, index: int) -> Format:
        cellxfs = self.get_cellXfs(index)
        numfmt = self.get_numfmt(int(cellxfs.numFmtId))
        font = self.get_fontid(int(cellxfs.fontId) - 1)
        fill = self.get_fillid(int(cellxfs.fillId))
        border = self.get_borders(int(cellxfs.borderId))

        return Format(numFmt=numfmt, font=font, fill=fill, border=border)
