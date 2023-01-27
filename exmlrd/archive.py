import os
import re
import xml.etree.ElementTree as ET
from enum import Enum
from typing import Any, Generator, List, Tuple, Union
from xml.etree.ElementTree import Element
from zipfile import ZipFile

from pydantic import validate_arguments

from exmlrd.cell import Cell, CellAttribute, Coordinate, Deco
from exmlrd.excel import ExcelObj


class XmlTag(Enum):
    DIMENSION = "dimension"
    SHEETVIEWS = "sheetViews"
    COLS = "cols"
    SHEETDATA = "sheetData"
    MERGECELLS = "mergeCells"
    PHONEITCPR = "phoneticPr"
    PAGEMARGINS = "pageMargins"
    PAGESETUP = "pageSetup"
    HEADERFOOTER = "headerFooter"
    DRAWING = "drawing"
    SHEETPR = "sheetPr"


class WorkbookTag(Enum):
    WORKBOOKPR = "workbookPr"
    WORKBOOKPROTECTION = "workbookProtection"
    BOOKVIEWS = "bookViews"
    SHEETS = "sheets"
    DEFINEDNAMES = "definedNames"
    CALCPR = "calcPr"


## Tags Name
# STRING = "s"
# FORMULA = "f"  # Calculations and evaluation formulas
# TYPE_FORMULA_CACHE_STRING = "str"
# ADRESS = "r"


class ExcelArchive:

    chart = r"xl/charts/chart.*.xml"
    drawings = r"xl/drawings/drawings.*.xml"
    printerSettings = r"xl/printerSettings/printerSettings.*.bin"
    theme = "xl/theme/theme.xml"
    worksheets = "xl/worksheets/sheet"
    styles = "xl/styles.xml"
    calcChain = "xl/calcChain.xml"
    sharedStrings = "xl/sharedStrings.xml"
    workbook = "xl/workbook.xml"

    def __init__(self, filepath: str):
        self.excel = ExcelObj(path=filepath)
        self.archive = self.__arch(self.excel.path)
        self.__worksheet = self.__get_worksheet()

    def __arch(self, path: str):
        return ZipFile(path)

    def get_archive_filename_all(self) -> Generator[str, None, None]:
        for info in self.archive.infolist():
            yield info.filename

    def get_archive_filename(self, filename: str):
        return [
            af.filename for af in self.archive.infolist() if af.filename == filename
        ]

    @validate_arguments
    def worksheet(self, index: int) -> str:
        if 0 > index and len(self.__worksheet) <= index:
            raise Exception(
                "The sheet could not be read. (Either the index is wrong or the sheet does not exist)"
            )
        return self.__worksheet[index - 1]

    def __get_elem_index(self, elem: Element, word: str) -> int:
        if not isinstance(elem, Element):
            return 0
        else:
            for index, __e in enumerate(elem):
                if self.del_namespace(__e.tag) == word:
                    return index
        return 0

    def __get_worksheet(self) -> List[Union[str, Any]]:
        f = self.workbook
        tree = ET.parse(self.archive.open(f))
        root_elem = tree.getroot()
        __eidx = self.__get_elem_index(root_elem, WorkbookTag.SHEETS.value)

        sheet_names = []
        for elem in root_elem[__eidx]:
            if elem.attrib.get("name"):
                sheet_names.append(elem.attrib["name"])
        return sheet_names

    def __worksheetpath(self, index: int) -> Union[str, List]:
        return self.worksheets + str(index) + ".xml"

    def get_cell(
        self, row: int, col: int, *, worksheet: int = 1
    ) -> Tuple[Coordinate, Cell, CellAttribute]:

        _coordinate = Coordinate(row=row, col=col)
        _prop = CellAttribute()

        f = self.__worksheetpath(worksheet)
        tree = ET.parse(self.archive.open(f))
        root_elem = tree.getroot()
        try:
            __eidx = self.__get_elem_index(root_elem, XmlTag.SHEETDATA.value)
            _ex_row = root_elem[__eidx][row][col]
        except IndexError as e:
            print(f"[cell({row},{col})] Warning: {e}")
            _cell = Cell()
            return _coordinate, _cell, _prop

        _cell = Cell(**_ex_row.attrib)
        _coordinate.address = _cell.r
        if _ex_row.attrib.get("t"):
            if _ex_row.attrib["t"] == "s":
                for _e_v in _ex_row:
                    if self.del_namespace(_e_v.tag) == "v":
                        _cell.v = _e_v.text if isinstance(_e_v.text, str) else ""
                        _prop = self.get_property(v_idx=_cell.v)
            elif _ex_row.attrib["t"] == "str":
                for _e_v in _ex_row:
                    if self.del_namespace(_e_v.tag) == "f":
                        _cell.f = _e_v.text if isinstance(_e_v.text, str) else ""
                    elif self.del_namespace(_e_v.tag) == "v":
                        _cell.v = _e_v.text if isinstance(_e_v.text, str) else ""
        else:
            for _e_v in _ex_row:
                if self.del_namespace(_e_v.tag) == "f":
                    _cell.f = _e_v.text if isinstance(_e_v.text, str) else ""
                elif self.del_namespace(_e_v.tag) == "v":
                    _cell.v = _e_v.text if isinstance(_e_v.text, str) else ""

        if _coordinate.address:
            _cell.m = self.get_mergecell(_coordinate.address, worksheet)

        if not _cell.bf:
            _cell.bf = self.get_basefont(worksheet)

        return _coordinate, _cell, _prop

    def get_property(self, *, v_idx: str = "1"):

        _attrib = CellAttribute(id=int(v_idx))
        _deco = Deco()

        f = self.sharedStrings
        tree = ET.parse(self.archive.open(f))
        root_elem = tree.getroot()

        for _e in root_elem[int(v_idx)]:
            _parents_tag = self.del_namespace(_e.tag)
            if _parents_tag == "t":
                """tag<t>: text"""
                _deco.text = _e.text if isinstance(_e.text, str) else ""
                _attrib.decorator.append(_deco)
            elif _parents_tag == "r":
                """tag<r>: rich text"""
                for _sub_e in _e:
                    if self.del_namespace(_sub_e.tag) == "t":
                        _deco.text = _sub_e.text if isinstance(_sub_e.text, str) else ""
                    elif self.del_namespace(_sub_e.tag) == "rPr":
                        for _eprop in _sub_e:
                            _prop_tag = self.del_namespace(_eprop.tag)
                            if _prop_tag == "sz":
                                _deco.sz = _eprop.attrib["val"]
                            elif _prop_tag == "color":
                                _deco.color = _eprop.attrib["theme"]
                            elif _prop_tag == "rFont":
                                _deco.rFont = _eprop.attrib["val"]
                            elif _prop_tag == "family":
                                _deco.family = _eprop.attrib["val"]
                            elif _prop_tag == "charset":
                                _deco.charset = _eprop.attrib["val"]
                            elif _prop_tag == "scheme":
                                _deco.scheme = _eprop.attrib["val"]
                            elif _prop_tag == "vertAlign":
                                _deco.vertAlign = _eprop.attrib["val"]
                _attrib.decorator.append(_deco)
                _deco = Deco()
        return _attrib

    def get_mergecell(self, start_cell: str, worksheet: int) -> str:
        f = self.__worksheetpath(worksheet)
        tree = ET.parse(self.archive.open(f))
        root_elem = tree.getroot()
        __eidx = self.__get_elem_index(root_elem, XmlTag.MERGECELLS.value)
        for elem in root_elem[__eidx]:
            if elem.attrib.get("ref"):
                if start_cell in elem.attrib["ref"].split(":"):
                    return elem.attrib["ref"]
        return ""

    def get_basefont(self, worksheet: int):
        f = self.__worksheetpath(worksheet)
        tree = ET.parse(self.archive.open(f))
        root_elem = tree.getroot()
        try:
            __eidx = self.__get_elem_index(root_elem, XmlTag.HEADERFOOTER.value)
            base = root_elem[__eidx][0].text
            if not isinstance(base, str):
                base = ""
        except IndexError:
            return ""

        m = re.match(r"^.*\"([\w].*)\"", base)
        if isinstance(m, str):
            font = m.group(1)
        else:
            font = base

        return str(font)

    def del_namespace(
        self,
        tag: str,
        namespace: str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    ):
        return str(tag).replace("{%s}" % namespace, "")
