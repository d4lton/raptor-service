#
# Copyright ©2024 Dana Basken
#

from pydantic import BaseModel

class ExcelPoolTask(BaseModel):
    type: str
    group_id: str
    item_id: str
    data: dict
