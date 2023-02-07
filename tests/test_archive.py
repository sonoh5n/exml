import pytest

from exmlrd import log
from exmlrd.archive import ExcelArchive
from exmlrd.cell import Format
from exmlrd.sharedstyle import SiTag

logger = log.get_logger(__name__)


def test_get_archive_filename_all(setup_excel):
    pred = [
        '[Content_Types].xml',
        '_rels/.rels',
        'xl/styles.xml',
        'xl/workbook.xml',
        'docProps/app.xml',
        'docProps/core.xml',
        'xl/theme/theme1.xml',
        'xl/_rels/workbook.xml.rels',
        'xl/worksheets/sheet1.xml'
    ]
    archive = ExcelArchive("tests/sample.xlsx")
    result = [f for f in archive.get_archive_filename_all()]
    assert set(result) == set(pred)

@pytest.mark.parametrize('arg, expect', [
    ('[Content_Types].xml', ['[Content_Types].xml']),
    ('_rels/.rels', ['_rels/.rels']),
    ('docProps/app.xml', ['docProps/app.xml']),
    ('xl/_rels/workbook.xml.rels', ['xl/_rels/workbook.xml.rels'])])
def test_get_archive_filename(setup_excel, arg, expect):
    archive = ExcelArchive("tests/sample.xlsx")
    print(arg)
    print(expect)
    assert archive.get_archive_filename("[Content_Types].xml") == ["[Content_Types].xml"]

def test_worksheet(setup_excel):
    archive = ExcelArchive("tests/sample.xlsx")
    ws = archive.worksheet(1)
    assert ws == "Test1"

@pytest.mark.parametrize('arg1, arg2, arg3', [
    (0, 0, 'A1'),
    (2, 3, 'D3'),
    (3, 1, 'B4')])
def test_get_celladdress(setup_excel, arg1, arg2, arg3):
    archive = ExcelArchive("tests/sample.xlsx")
    cell = archive.get_cell(arg1,arg2)
    assert cell.row == arg1
    assert cell.col == arg2
    assert cell.address == arg3


@pytest.mark.parametrize('row, col, arg1, arg2, arg3, arg4, arg5, arg6, arg7', 
    [(1, 1, 1, 1, "B2", "57","", SiTag(), Format()),
    (2, 1, 2, 1, "B3", "19", "", SiTag(), Format())])
def test_get_cells(setup_excel, row, col, arg1, arg2, arg3, arg4, arg5, arg6, arg7):
    archive = ExcelArchive("tests/sample.xlsx")
    cell = archive.get_cell(row,col)
    logger.debug(cell)
    assert cell.row == arg1
    assert cell.col == arg2
    assert cell.address == arg3
    assert cell.value == arg4
    assert cell.formula == arg5
    assert cell.shared == arg6
    assert cell.style != arg7

@pytest.mark.parametrize('start_cell, sheet, pred', [
    ("A1", 1, ""),("H1", 1, "H1:I1")])
def test_get_mergecell(setup_excel, start_cell, sheet, pred):
    archive = ExcelArchive("tests/sample.xlsx")
    result = archive.get_mergecell(start_cell, sheet)
    assert result == pred

def test_get_all_mergecell(setup_excel):
    pred = {"1": ["H1:I1", 'A7:G9']}
    archive = ExcelArchive("tests/sample.xlsx")
    result = archive.get_all_mergecell(1)
    assert list(result.keys()) == ["1"]
    assert set(result.get("1")) == set(pred.get("1"))
