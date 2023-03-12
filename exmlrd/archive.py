import json
from dataclasses import asdict
from typing import Any, Generator, Optional
from zipfile import ZipFile

from pydantic import validate_arguments

from exmlrd.cell import Cell, SheetXml
from exmlrd.chart import XmlChart
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
        self.chart = XmlChart(self.archive)
        self.chart.get_chart(2)
        self.sheetnum = 1

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
        if ws:
            self.sheetnum = index
        return ws

    def get_cell(self, row: int, col: int, *, worksheet: int = 1) -> Cell:
        __cell = self.sheetxml.get_cell(row, col, worksheet=worksheet)
        return __cell

    def get_mergecell(self, start_cell: str, worksheet: int) -> str:
        __merge_cell = self.sheetxml.get_mergecell(start_cell, worksheet)
        return __merge_cell

    def get_all_mergecell(self, worksheet: int) -> dict[str, list[str]]:
        __merge_cells = self.sheetxml.get_mergecells(worksheet)
        return __merge_cells.ref

    def get_sheetrange_address(self) -> Optional[str]:
        return self.sheetxml.get_dimension_address(worksheet=self.sheetnum)

    def get_sheetrange_coordinate(self, *, sheetnum=None):
        if sheetnum is None:
            sheetnum = self.sheetnum
        return self.sheetxml.get_dimension_coordinate(worksheet=sheetnum)

    def get_charts_list(self):
        raise RuntimeError

    def get_chart(self, index: int):
        raise RuntimeError

    def to_json(
        self,
        *,
        row: Optional[int] = None,
        col: Optional[int] = None,
        save_path: Optional[str] = None
    ):
        ws_name = self.worksheet(self.sheetnum)
        contents: dict[str, list[Any]] = {ws_name: []}
        if (row is None) or (col is None):
            start_coordinate, end_coordinate = self.get_sheetrange_coordinate()
            for r in range(start_coordinate[0], end_coordinate[0]):
                for c in range(start_coordinate[1], end_coordinate[1]):
                    cell = self.get_cell(r + 1, c + 1, worksheet=self.sheetnum)
                    contents[ws_name].append(asdict(cell))
        else:
            end_coordinate = (row, col)
            cell = self.get_cell(row, col, worksheet=self.sheetnum)
            contents[ws_name].append(asdict(cell))

        if save_path:
            with open(save_path, mode="w", encoding="utf-8") as f:
                json.dump(contents, f, indent=2, ensure_ascii=False)

        return json.dumps(contents, indent=2, ensure_ascii=False)
