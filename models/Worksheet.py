#
# Copyright ©2024 Dana Basken
#

import re
import numpy

class Worksheet:

    def __init__(self, worksheet_name: str, workbook):
        self.name = worksheet_name
        self._workbook = workbook
        self._worksheet = self._workbook.Worksheets[self.name]

    def get_durable_ids(self):
        results = {}
        rows = self.get_used_range_values()
        for row in rows:
            durable_id, durable_id_type = self.get_durable_id(row)
            if durable_id_type is not None:
                results[durable_id] = self.get_durable_id_values(row, durable_id_type)
        return results

    @staticmethod
    def get_durable_id(row):
        durable_id = row[1]
        if durable_id is None: return None, None
        match = re.search(r"^(\w+?)___(.+)$|^(\w+?)\.(.+)$", durable_id)
        if match.group(1) == "settings": return durable_id, "SETTING"
        if match.group(1) == "metadata": return durable_id, "METADATA"
        return durable_id, "METRIC"

    @staticmethod
    def get_durable_id_values(row, durable_id_type):
        match durable_id_type:
            case "METRIC":
                values = numpy.array(row)
                return values[10:]
            case _:
                return row[3]

    def get_used_range_values(self):
        used_range = self._worksheet.UsedRange
        row_count = used_range.Rows.Count + 1
        column_count = used_range.Columns.Count + 1
        if column_count > 1000: column_count = 1001
        value_range = self._worksheet.Range(used_range.Cells(1, 1), used_range.Cells(row_count, column_count))
        return value_range.Value
