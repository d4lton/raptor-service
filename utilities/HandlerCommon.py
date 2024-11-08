#
# Copyright ©2024 Dana Basken
#

from tenacity import retry, stop_after_attempt, wait_random_exponential
import exceltypes
from models.Worksheet import Worksheet
from models.DurableIds import DurableIds

# these are "retry" versions of win32com/Excel-related functions that might fail due to single-threaded nature of Excel

@retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
def open_workbook(excel: exceltypes.Application, file_path: str) -> exceltypes.Workbook:
    return excel.Workbooks.Open(file_path)

@retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
def get_durable_ids(worksheet: Worksheet) -> DurableIds:
    return worksheet.get_durable_ids()

@retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
def set_durable_id_values(durable_ids: DurableIds, durable_id: str, values: any):
    durable_ids.set_durable_id_values(durable_id, values)

@retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
def calculate(workbook: exceltypes.Workbook):
    workbook.Application.CalculateFull()
