import os

from pydantic import BaseModel, validator

from exmlrd.exceptions import NotSupportFmt


class ExcelObj(BaseModel):

    path: str

    @validator("path", always=True)
    def check_ext(cls, v):
        _, ext = os.path.splitext(v)
        if ext == ".xlsx":
            ...
        else:
            raise NotSupportFmt(f"not support format:{v}")
        return v

    @validator("path", always=True)
    def check_exist(cls, v):
        if not os.path.exists(v):
            raise FileNotFoundError(f"File not exist: {v}")
        return v
