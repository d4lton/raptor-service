#
# Copyright ©2024 Dana Basken
#

import re
import exceltypes
import numpy

class DurableIds:

    def __init__(self, worksheet: exceltypes.Worksheet, max_columns: int = 1000):
        self._worksheet = worksheet
        self._max_columns = max_columns
        self._durable_id_values: dict[str, numpy.ndarray] = {}
        self._durable_id_types: dict[str, str] = {}
        self._build_durable_ids()

    def get_durable_id(self, durable_id: str) -> tuple[numpy.ndarray, str]:
        return self._durable_id_values[durable_id], self._durable_id_types[durable_id]

    def set_durable_id_values(self, durable_id: str, values: numpy.ndarray):
        self._durable_id_values[durable_id] = values
        # TODO: set these values back into self._worksheet at the correct row and columns

    def get_durable_ids(self) -> tuple[dict[str, str], dict[str, numpy.ndarray]]:
        return self._durable_id_types, self._durable_id_values

    def _build_durable_ids(self):
        range = self._get_used_range_with_constraint()
        rows = range.Value
        for row in rows:
            durable_id, durable_id_type = self._get_durable_id(row)
            if durable_id_type is not None:
                self._store_durable_id(durable_id, durable_id_type, self._get_durable_id_values(row, durable_id_type))

    @staticmethod
    def _get_durable_id(row) -> tuple[str | None, str | None]:
        durable_id = row[1]
        if durable_id is None: return None, None
        match = re.search(r"^(\w+?)___(.+)$|^(\w+?)\.(.+)$", durable_id)
        if match is not None:
            if match.group(1) is not None:
                prefix = match.group(1)
                id = match.group(2)
            else:
                prefix = match.group(3)
                id = match.group(4)
            if prefix == "settings": return id, "SETTING"
            if prefix == "metadata": return id, "METADATA"
            return durable_id, "METRIC"
        return None, None

    @staticmethod
    def _get_durable_id_values(row, durable_id_type) -> numpy.ndarray | float:
        match durable_id_type:
            case "METRIC":
                values = numpy.array(row)[10:]
                while values.size > 0 and values[-1] is None: values = values[:-1] # trim None elements from end of array
                return values
            case _:
                return row[3] # TODO: assert this is a float?

    def _store_durable_id(self, durable_id, durable_id_type, values):
        self._durable_id_types[durable_id] = durable_id_type
        self._durable_id_values[durable_id] = values

    def _get_used_range_with_constraint(self) -> exceltypes.Range:
        used_range = self._worksheet.UsedRange
        row_count = used_range.Rows.Count + 1
        column_count = used_range.Columns.Count + 1
        if column_count > self._max_columns: column_count = self._max_columns + 1
        return self._worksheet.Range(used_range.Cells(1, 1), used_range.Cells(row_count, column_count))
