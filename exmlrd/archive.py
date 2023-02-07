"""
Copyright (c) 2023 HAYATO SONOKAWA
"""

from typing import Dict, Generator, List, Optional
from zipfile import ZipFile

from pydantic import validate_arguments

from exmlrd.cell import Cell, SheetXml
from exmlrd.excel import ExcelObj


class ExcelArchive:
    """
    Obtain archive information such as the xml files that make up Excel.

    Mainly, sheetX.xml, styles.xml, workbook.xml, and sharedStrings.xml are acquired and analyzed.

    Attributes:
        excel (ExcelObj): Excel file to be read
        archive (ZipFile): Excel archive information
        sheetxml (SheetXml): Object with information parsed from sheetX.xml
    """

    def __init__(self, filepath: str):
        self.excel = ExcelObj(path=filepath)
        self.archive = self.__arch(self.excel.path)
        self.sheetxml = SheetXml(self.archive)

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
        ws = self.sheetxml.worksheet(index)
        return ws

    def get_cell(self, row: int, col: int, *, worksheet: int = 1) -> Optional[Cell]:
        __cell = self.sheetxml.get_cell(row, col, worksheet=worksheet)
        return __cell

    def get_mergecell(self, start_cell: str, worksheet: int) -> str:
        __merge_cell = self.sheetxml.get_mergecell(start_cell, worksheet)
        return __merge_cell

    def get_all_mergecell(self, worksheet: int) -> dict[str, list[str]]:
        __merge_cells = self.sheetxml.get_mergecells(worksheet)
        return __merge_cells.ref
