#
# Copyright ©2024 Dana Basken
#

import re
import numpy
import pandas

class DurableIds:

    def __init__(self, worksheet, max_columns: int = 1000):
        self._worksheet = worksheet
        self._max_columns = max_columns
        self._durable_id_values = {}
        self._durable_id_types = {}
        self._build_durable_ids()

    def get_durable_id(self, durable_id: str):
        return self._durable_id_values[durable_id], self._durable_id_types[durable_id]

    def set_durable_id_values(self, durable_id: str, values: numpy.ndarray):
        self._durable_id_values[durable_id] = values

    def get_durable_ids(self):
        return self._durable_id_types, self._durable_id_values

    def _build_durable_ids(self):
        range = self._get_used_range_with_constraint()
        rows = range.Value
        for row in rows:
            durable_id, durable_id_type = self._get_durable_id(row)
            if durable_id_type is not None:
                self._set_durable_id(durable_id, durable_id_type, self._get_durable_id_values(row, durable_id_type))

    @staticmethod
    def _get_durable_id(row):
        durable_id = row[1]
        if durable_id is None: return None, None
        match = re.search(r"^(\w+?)___(.+)$|^(\w+?)\.(.+)$", durable_id)
        if match.group(1) == "settings": return durable_id, "SETTING"
        if match.group(1) == "metadata": return durable_id, "METADATA"
        return durable_id, "METRIC"

    @staticmethod
    def _get_durable_id_values(row, durable_id_type):
        match durable_id_type:
            case "METRIC":
                values = numpy.array(row)[10:]
                return numpy.nan_to_num(pandas.to_numeric(values, errors="coerce"))
            case _:
                return row[3]

    def _set_durable_id(self, durable_id, durable_id_type, values):
        self._durable_id_types[durable_id] = durable_id_type
        self._durable_id_values[durable_id] = values

    def _get_used_range_with_constraint(self):
        used_range = self._worksheet.UsedRange
        row_count = used_range.Rows.Count + 1
        column_count = used_range.Columns.Count + 1
        if column_count > self._max_columns: column_count = self._max_columns + 1
        return self._worksheet.Range(used_range.Cells(1, 1), used_range.Cells(row_count, column_count))
