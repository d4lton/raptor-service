#
# Copyright ©2024 Dana Basken
#

from models.DurableIds import DurableIds


class Worksheet:

    def __init__(self, worksheet_name: str, workbook: any, max_columns: int = 1000):
        self.name = worksheet_name
        self._max_columns = max_columns
        self._workbook = workbook
        self._worksheet = self._workbook.Worksheets[self.name]

    def get_durable_ids(self):
        return DurableIds(self._worksheet, self._max_columns)
