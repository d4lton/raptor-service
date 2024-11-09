#
# Copyright ©2024 Dana Basken
#

from task_handlers.BaseHandler import BaseHandler
from models.Worksheet import Worksheet

class DemoTaskHandler(BaseHandler):

    def __init__(self):
        super().__init__()

    def process(self):
        worksheet = Worksheet("M - Monthly", self.workbook)
        self.add_response("running", "get_durable_ids")
        durable_ids = self.get_durable_ids(worksheet)
        self.add_response("running", "get_durable_id")
        income_statement_returns, durable_id_type = durable_ids.get_durable_id("incomeStatement.returns")
        # TODO: do something to income_statement_returns
        self.add_response("running", "set_durable_id_values")
        self.set_durable_id_values(durable_ids, "incomeStatement.returns", income_statement_returns)
        self.add_response("running", "calculate_workbook")
        self.calculate(self.workbook)
        self.add_response("running", "get_durable_ids_2")
        durable_ids = self.get_durable_ids(worksheet)
        # TODO: send durable_ids to DWH
