#
# Copyright ©2024 Dana Basken
#

from pydantic import BaseModel
from typing import Literal

ExcelPoolTaskType = Literal["demo"]

class ExcelPoolTask(BaseModel):
    type: ExcelPoolTaskType
    site_id: str
    item_id: str
    data: dict
