import pytest

from exmlrd.archive import ExcelArchive, XmlTag


def test_xmltag():
    assert XmlTag.DIMENSION.value == "dimension"
    assert XmlTag.SHEETVIEWS.value == "sheetViews"
    assert XmlTag.COLS.value == "cols"
    assert XmlTag.SHEETDATA.value == "sheetData"
    assert XmlTag.MERGECELLS.value == "mergeCells"
    assert XmlTag.PHONEITCPR.value == "phoneticPr"
    assert XmlTag.PAGEMARGINS.value == "pageMargins"
    assert XmlTag.PAGESETUP.value == "pageSetup"
    assert XmlTag.HEADERFOOTER.value == "headerFooter"
    assert XmlTag.DRAWING.value == "drawing"
    assert XmlTag.SHEETPR.value == "sheetPr"


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
def test_get_cell(setup_excel, arg1, arg2, arg3):
    archive = ExcelArchive("tests/sample.xlsx")
    cell, _, _ = archive.get_cell(arg1,arg2)
    assert cell.row == arg1
    assert cell.col == arg2
    assert cell.address == arg3


@pytest.mark.parametrize('row, col, arg1, arg2, arg3, arg4, arg5, arg6, arg7', 
    [(1, 1, "B2", "1", "n", "","","",""),
    (2, 1, "B3", "1", "n", "","","","")])
def test_get_attr(setup_excel, row, col, arg1, arg2, arg3, arg4, arg5, arg6, arg7):
    archive = ExcelArchive("tests/sample.xlsx")
    _, attr, _ = archive.get_cell(row,col)
    assert attr.r == arg1
    assert attr.s == arg2
    assert attr.t == arg3
    assert attr.v == arg4
    assert attr.f == arg5
    assert attr.m == arg6
    assert attr.bf == arg7

@pytest.mark.parametrize('row, col, arg1', [("A1", 1, ""),("H1", 1, "H1:I1")])
def test_get_mergecell(setup_excel, row, col, arg1):
    archive = ExcelArchive("tests/sample.xlsx")
    result = archive.get_mergecell(row, col)
    assert result == arg1

def test_get_basefont(setup_excel):
    archive = ExcelArchive("tests/sample.xlsx")
    result = archive.get_basefont(1)
    assert result == ""

def test_del_namespace(setup_excel):
    archive = ExcelArchive("tests/sample.xlsx")
    n = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}scheme"
    result = archive.del_namespace("scheme", namespace=n)
    assert result == "scheme"
    n = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}/test/scheme"
    result = archive.del_namespace("/test/scheme", namespace=n)