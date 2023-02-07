from enum import Enum


class SheetXmlTag(Enum):
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


class StylesTag(Enum):
    FONTS = "fonts"
    FILLS = "fills"
    BORDERS = "borders"
    CELLSTYLEXFS = "cellStyleXfs"
    CELLXFS = "cellXfs"
    CELLSTYLES = "cellStyles"
    DXFS = "dxfs"
    TABLESTYLES = "tableStyles"
    EXTLST = "extLst"
    NUMFMTS = "numFmts"


class SharedItemTag(Enum):
    SI = "si"
    R = "r"
    RPR = "rPr"
    PHONETICPR = "phoneticPr"
