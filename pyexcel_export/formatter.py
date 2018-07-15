import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.copier import WorksheetCopy

import os
import logging
import json
from io import BytesIO
from collections import OrderedDict

from .serialize import MyEncoder

debug_logger = logging.getLogger('debug')


class ExcelFormatter:
    def __init__(self, template_file: str=None):
        if template_file and os.path.exists(template_file):
            self.styled_wb = openpyxl.load_workbook(template_file)
            self.to_stylesheets(self.styled_wb)
        else:
            self.styled_wb = openpyxl.Workbook()
            self.styled_wb.active.title = '_template'

    @property
    def data(self):
        output = BytesIO()
        self.styled_wb.save(output)

        return output

    @data.setter
    def data(self, _styles):
        self.styled_wb = openpyxl.load_workbook(_styles)

    def save(self, raw_data, out_file, meta=None, retain_meta=True):
        if '_styles' in meta.keys():
            self.data = meta['_styles']
            meta['_styles'] = self.data

        if os.path.exists(out_file):
            wb = openpyxl.load_workbook(out_file)
            self.to_stylesheets(wb)
            self.append_styled_sheets(wb)

            original_sheet_names = []
            extraneous_sheet_names = []
        else:
            wb = self.styled_wb
            original_sheet_names = wb.sheetnames
            extraneous_sheet_names = []

        extraneous_sheet_names.append('_template')
        inserted_sheets = []

        if not retain_meta:
            extraneous_sheet_names.append('_meta')
        else:
            if '_meta' in original_sheet_names:
                wb.remove(wb['_meta'])

            self.create_styled_sheet(wb, '_meta', 0)

            meta_matrix = []

            for k, v in meta.excel_view.items():
                if not k.startswith('_'):
                    if isinstance(v, (dict, OrderedDict)):
                        v = json.dumps(v, cls=MyEncoder)
                    meta_matrix.append([k, v])

            self.fill_matrix(wb['_meta'], meta_matrix, rules=meta)

            if '_hidden' in meta.keys():
                starting_index = len(meta_matrix) + 1

                for k, v in meta['_hidden'].items():
                    wb['_meta'].append()

                    table_matrix = [['# ' + k]]
                    table_matrix.extend(v)
                    self.fill_matrix(wb['_meta'], table_matrix, starting_index, rules=meta)

                    starting_index += len(table_matrix) + 1

        if '_meta' in raw_data.keys():
            raw_data.pop('_meta')

        for sheet_name, cell_matrix in raw_data.items():
            if sheet_name not in original_sheet_names:
                self.create_styled_sheet(wb, sheet_name)
                inserted_sheets.append(sheet_name)
            else:
                self.fill_matrix(wb[sheet_name], cell_matrix, rules=meta)

            ws = wb[sheet_name]
            for row_num, row in enumerate(cell_matrix):
                for col_num, value in enumerate(row):
                    if isinstance(value, (dict, OrderedDict)):
                        value = json.dumps(value, cls=MyEncoder)

                    ws.cell(column=(col_num + 1),
                            row=(row_num + 1),
                            value=value)

        for sheet_name in extraneous_sheet_names:
            if sheet_name in wb.sheetnames:
                wb.remove(wb[sheet_name])

        for sheet_name in wb.sheetnames:
            if sheet_name.startswith('_') and sheet_name != '_meta':
                wb.remove(wb[sheet_name])

        wb.save(out_file)

    @staticmethod
    def create_styled_sheet(wb, sheet_name, pos: int=None):
        wb.create_sheet(sheet_name, pos)
        if '_template' in wb.sheetnames and sheet_name != '_template':
            WorksheetCopy(wb['_template'], wb[sheet_name]).copy_worksheet()
            # wb[sheet_name].copy_worksheet(wb['_template'])

    def append_styled_sheets(self, wb):
        if '_template' not in wb.sheetnames:
            wb.create_sheet('_template')
            if '_template' in self.styled_wb.sheetnames:
                WorksheetCopy(self.styled_wb['_template'], wb['_template']).copy_worksheet()
                # wb['_template'].copy_worksheet(wb['_template'])

        return wb

    @staticmethod
    def to_stylesheets(wb):
        for ws in wb:
            for row in ws:
                for cell in row:
                    cell.value = None

        return wb

    @staticmethod
    def fill_matrix(ws, cell_matrix, start_row=0, rules=None):
        for row_num, row in enumerate(cell_matrix):
            for col_num, value in enumerate(row):
                if isinstance(value, (dict, OrderedDict)):
                    value = json.dumps(value, cls=MyEncoder)

                ws.cell(column=(col_num + 1),
                        row=(row_num + start_row + 1),
                        value=value)

        if rules is not None:
            if rules.get('has_header', False) and rules.get('freeze_header', False):
                if ws.title != '_meta':
                    ws.freeze_panes = 'A2'
            if rules.get('col_width_fit_param_keys', False):
                width = max([len(str(cell.value)) for cell in next(ws.iter_cols())])
                ws.column_dimensions['A'].width = width + 2
            if rules.get('col_width_fit_ids', False):
                for i, header_cell in enumerate(next(ws.iter_rows())):
                    header_item = header_cell.value
                    if header_item and header_item.endswith('id'):
                        col_letter = get_column_letter(i + 1)
                        width = max([len(str(cell.value)) for cell in list(ws.iter_cols())[i]])
                        ws.column_dimensions[col_letter].width = width + 2
