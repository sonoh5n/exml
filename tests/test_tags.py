import pytest

from exmlrd.tags import SharedItemTag, SheetXmlTag, StylesTag, WorkbookTag


def test_SheetXmlTag():
    assert SheetXmlTag.DIMENSION.value == "dimension"
    assert SheetXmlTag.SHEETVIEWS.value == "sheetViews"
    assert SheetXmlTag.COLS.value == "cols"
    assert SheetXmlTag.SHEETDATA.value == "sheetData"
    assert SheetXmlTag.MERGECELLS.value == "mergeCells"
    assert SheetXmlTag.PHONEITCPR.value == "phoneticPr"
    assert SheetXmlTag.PAGEMARGINS.value == "pageMargins"
    assert SheetXmlTag.PAGESETUP.value == "pageSetup"
    assert SheetXmlTag.HEADERFOOTER.value == "headerFooter"
    assert SheetXmlTag.DRAWING.value == "drawing"
    assert SheetXmlTag.SHEETPR.value == "sheetPr"

def test_SharedItemTag():
    assert SharedItemTag.SI.value == "si"
    assert SharedItemTag.R.value == "r"
    assert SharedItemTag.RPR.value == "rPr"
    assert SharedItemTag.PHONETICPR.value == "phoneticPr"

def test_WorkbookTag():
    assert WorkbookTag.WORKBOOKPR.value == "workbookPr"
    assert WorkbookTag.WORKBOOKPROTECTION.value == "workbookProtection"
    assert WorkbookTag.BOOKVIEWS.value == "bookViews"
    assert WorkbookTag.SHEETS.value == "sheets"
    assert WorkbookTag.DEFINEDNAMES.value == "definedNames"
    assert WorkbookTag.CALCPR.value == "calcPr"

def test_StylesTag():
    assert StylesTag.FONTS.value == "fonts" 
    assert StylesTag.FILLS.value == "fills" 
    assert StylesTag.BORDERS.value == "borders" 
    assert StylesTag.CELLSTYLEXFS.value == "cellStyleXfs" 
    assert StylesTag.CELLXFS.value == "cellXfs" 
    assert StylesTag.CELLSTYLES.value == "cellStyles" 
    assert StylesTag.DXFS.value == "dxfs" 
    assert StylesTag.TABLESTYLES.value == "tableStyles" 
    assert StylesTag.EXTLST.value == "extLst" 
    assert StylesTag.NUMFMTS.value == "numFmts"
