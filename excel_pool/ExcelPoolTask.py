#
# Copyright ©2024 Dana Basken
#

from pydantic import BaseModel

class ExcelPoolTask(BaseModel):
    site_id: str
    item_id: str
