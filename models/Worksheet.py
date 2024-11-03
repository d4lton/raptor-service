#
# Copyright ©2024 Dana Basken
#

import exceltypes
from models.DurableIds import DurableIds

class Worksheet:

    def __init__(self, worksheet_name: str, workbook: exceltypes.Workbook, max_columns: int = 1000):
        self.name = worksheet_name
        self._max_columns = max_columns
        self._workbook = workbook
        self._worksheet: exceltypes.Worksheet = self._workbook.Worksheets[self.name]

    def get_durable_ids(self) -> DurableIds:
        return DurableIds(self._worksheet, self._max_columns)
