"""
Copyright (c) 2023 HAYATO SONOKAWA
"""
import xml.etree.ElementTree as ET
from functools import cache
from typing import Dict, List, Optional
from xml.etree.ElementTree import Element
from zipfile import ZipFile

from pydantic import Field, dataclasses, validate_arguments, validator

from exmlrd import log
from exmlrd.exceptions import CellOutsideRange, NotFoundSheet
from exmlrd.sharedstyle import SharedStyle, SiTag
from exmlrd.styles import Format, Styels
from exmlrd.tags import SheetXmlTag, WorkbookTag
from exmlrd.tools import del_namespace

logger = log.get_logger(__name__)


@dataclasses.dataclass
class Cell:
    row: int = 0
    col: int = 0
    address: str = ""
    value: str = ""
    formula: str = ""
    shared: Optional[SiTag] = None
    style: Optional[Format] = None

    @validator("row", always=True)
    def check_row_range(cls, v):
        if 0 <= v and v <= 1048575:
            return v
        else:
            raise CellOutsideRange(
                "Invalid row number. The maximum and minimum values for the row are exceeded."
            )

    @validator("col", always=True)
    def check_col_range(cls, v):
        if 0 <= v and v <= 18278:
            return v
        else:
            raise CellOutsideRange(
                "Invalid col number. The maximum and minimum values for the col are exceeded."
            )


@dataclasses.dataclass
class MergeCell:
    ref: dict[str, list[str]] = Field(default_factory=dict)


class Worksheet:
    """
    Class for retrieving and manipulating worksheet information

    Attributes:
        archive (ZipFile): Excel archive information
        worksheets (List[str]): Stores all worksheet names.
    """

    workbook_path = "xl/workbook.xml"
    worksheet_path = "xl/worksheets/sheet"

    def __init__(self, archive: ZipFile):
        self.archive = archive
        self.worksheets = self.__get_worksheet_name()

    def __get_elem_index(self, elem: Element, word: str) -> Optional[int]:
        if not isinstance(elem, Element):
            return None
        else:
            for index, __e in enumerate(elem):
                if del_namespace(__e.tag) == word:
                    return index
        return None

    @validate_arguments
    def get_worksheet(self, index: int) -> str:
        if len(self.worksheets) < index:
            raise NotFoundSheet(
                "The sheet could not be read. (Either the index is wrong or the sheet does not exist)"
            )
        return self.worksheets[index - 1]

    def __get_worksheet_name(self) -> List[str]:
        f = self.workbook_path
        tree = ET.parse(self.archive.open(f))
        root_elem = tree.getroot()
        __eidx = self.__get_elem_index(root_elem, WorkbookTag.SHEETS.value)
        if not isinstance(__eidx, int):
            return []

        sheet_names = []
        for elem in root_elem[__eidx]:
            if elem.attrib.get("name"):
                sheet_names.append(elem.attrib["name"])
        return sheet_names

    @validate_arguments
    def get_worksheetpath(self, index: int):
        return self.worksheet_path + str(index) + ".xml"


class SheetXml:
    """
    Class for acquiring and manipulating information parsed from sheetX.xml

    Attributes:
        archive (ZipFile): Excel archive information
        ws (Worksheet): Objects that manipulate worksheet information
        sharedstyle(SharedStyle): Objects that manipulate common font information
        style(Styels): Objects that manipulate font, border, and fill information
    """

    worksheets_basepath = "xl/worksheets/sheet"

    def __init__(self, archive: ZipFile):
        self.archive = archive
        self.ws = Worksheet(self.archive)
        self.sharedstyle = SharedStyle(self.archive)
        self.style = Styels(self.archive)

    def __get_elem_index(self, elem: Element, word: str) -> Optional[int]:
        if not isinstance(elem, Element):
            return None
        else:
            for index, __e in enumerate(elem):
                if del_namespace(__e.tag) == word:
                    return index
        return None

    def worksheet(self, index: int) -> str:
        return self.ws.get_worksheet(index)

    def get_cell(self, row: int, col: int, *, worksheet: int = 1):

        f = self.ws.get_worksheetpath(worksheet)
        tree = ET.parse(self.archive.open(f))
        root_elem = tree.getroot()
        __eidx = self.__get_elem_index(root_elem, SheetXmlTag.SHEETDATA.value)
        if not isinstance(__eidx, int):
            return None
        try:
            _ex_row = root_elem[__eidx][row][col]
        except IndexError as e:
            meg = f"[cell({row},{col})] Warning: {e}"
            logger.debug(meg)
            return None

        sidx = _ex_row.attrib.get("s")
        _cell = Cell(
            row=row,
            col=col,
            address=_ex_row.attrib["r"],
            style=self.style.get_format(int(str(sidx))),
        )
        if _ex_row.attrib.get("t"):
            if _ex_row.attrib["t"] == "s":
                for _e_v in _ex_row:
                    if del_namespace(_e_v.tag) == "v":
                        _cell.value = _e_v.text if isinstance(_e_v.text, str) else ""
                        __prop = self.sharedstyle.get_shareitem(
                            index=int(str(_e_v.text))
                        )
                        _cell.shared = __prop
                        _cell.value = __prop.text
            elif _ex_row.attrib["t"] == "str":
                for _e_v in _ex_row:
                    if del_namespace(_e_v.tag) == "f":
                        _cell.formula = _e_v.text if isinstance(_e_v.text, str) else ""
                    elif del_namespace(_e_v.tag) == "v":
                        _cell.value = _e_v.text if isinstance(_e_v.text, str) else ""
            elif _ex_row.attrib["t"] == "n":
                for _e_v in _ex_row:
                    if del_namespace(_e_v.tag) == "v":
                        _cell.value = _e_v.text if isinstance(_e_v.text, str) else ""
        else:
            for _e_v in _ex_row:
                if del_namespace(_e_v.tag) == "f":
                    _cell.formula = _e_v.text if isinstance(_e_v.text, str) else ""
                elif del_namespace(_e_v.tag) == "v":
                    _cell.value = _e_v.text if isinstance(_e_v.text, str) else ""

        return _cell

    @cache
    def get_mergecell(self, start_cell: str, worksheet: int) -> str:
        f = self.ws.get_worksheetpath(worksheet)
        tree = ET.parse(self.archive.open(f))
        root_elem = tree.getroot()
        __eidx = self.__get_elem_index(root_elem, SheetXmlTag.MERGECELLS.value)
        if not isinstance(__eidx, int):
            return ""
        for elem in root_elem[__eidx]:
            if elem.attrib.get("ref"):
                if start_cell in elem.attrib["ref"].split(":"):
                    return elem.attrib["ref"]
        return ""

    def get_mergecells(self, worksheet: int) -> MergeCell:
        mgcell = MergeCell()
        f = self.ws.get_worksheetpath(worksheet)
        tree = ET.parse(self.archive.open(f))
        root_elem = tree.getroot()
        __eidx = self.__get_elem_index(root_elem, SheetXmlTag.MERGECELLS.value)
        if not isinstance(__eidx, int):
            return mgcell

        for elem in root_elem[__eidx]:
            if elem.attrib.get("ref"):
                mgcell.ref.setdefault(str(worksheet), [])
                mgcell.ref[str(worksheet)].append((elem.attrib["ref"]))
        return mgcell
