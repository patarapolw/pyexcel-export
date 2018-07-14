import openpyxl

import os
from collections import OrderedDict


class ExcelFormatter:
    def __init__(self, template_file: str):
        self.template = OrderedDict()

        wb = openpyxl.load_workbook(template_file)
        for ws in wb:
            self.template[ws.title] = self.get_styles(ws)

    @staticmethod
    def get_styles(src):
        """

        :param openpyxl.worksheet.worksheet.Worksheet src:
        :param openpyxl.worksheet.worksheet.Worksheet dst:
        :return:

        TODO: load default formattings from a *.xlsx file
        TODO: verify that the styles are correctly formatted
        """
        styles = dict()

        # Get column_dimensions
        styles['column_dimensions'] = dict()
        for col_letter, col_data in src.column_dimensions.items():
            styles['column_dimensions'][col_letter] = col_data.__dict__

        return styles
