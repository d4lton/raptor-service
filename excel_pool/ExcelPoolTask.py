#
# Copyright ©2024 Dana Basken
#

from pydantic import BaseModel

class ExcelPoolTask(BaseModel):
    path: str