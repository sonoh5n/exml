from typing import List

from pydantic import Field, dataclasses, validator

from exmlrd.exceptions import CellOutsideRange


@dataclasses.dataclass
class Coordinate:
    row: int = 0
    col: int = 0
    address: str = ""

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
class Cell:
    r: str = ""
    s: str = ""
    t: str = ""
    v: str = ""
    f: str = ""
    m: str = ""
    bf: str = ""  # base font


@dataclasses.dataclass
class Deco:
    text: str = ""
    sz: str = ""
    color: str = ""
    rFont: str = ""
    family: str = ""
    charset: str = ""
    scheme: str = ""
    vertAlign: str = ""


@dataclasses.dataclass
class CellAttribute:
    id: int = 0
    decorator: List = Field(default_factory=list)
