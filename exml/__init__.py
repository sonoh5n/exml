from exml.archive import ExcelArchive


def excel_archiver(filepath: str) -> ExcelArchive:
    return ExcelArchive(filepath)
