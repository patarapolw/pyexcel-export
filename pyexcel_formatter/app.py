import pyexcel
from collections import OrderedDict
from datetime import datetime
import os
import json
import copy

from .serialize import RowExport, PyexcelExportEncoder, RowExportEncoder
from .defaults import DEFAULT_META


class ExcelFormatter:
    def __init__(self, in_file: str):
        self.in_file = in_file
        self.meta = DEFAULT_META

        in_base, in_format = os.path.splitext(in_file)

        if in_format == '.ods' or in_format == '.xlsx':
            self.data = self._load_pyexcel_excel()
        elif in_format == '.json':
            if os.path.splitext(in_base)[1] == '.pyexcel':
                self.data = self._load_pyexcel_json()
            else:
                self.data = self._load_json()
        else:
            raise ValueError('Unsupported file format, {}.'.format(in_format))

    def _load_pyexcel_excel(self):
        updated_data = pyexcel.get_book_dict(file_name=self.in_file)

        return self._set_updated_data(updated_data)

    def _load_pyexcel_json(self):
        with open(self.in_file) as f:
            data = json.load(f, object_pairs_hook=OrderedDict)

        for sheet_name, sheet_data in data.items():
            for j, row in enumerate(sheet_data):
                for i, cell in enumerate(row):
                    data[sheet_name][j][i] = json.loads(cell)

        return self._set_updated_data(data)

    def _load_json(self):
        with open(self.in_file) as f:
            data = json.load(f, object_pairs_hook=OrderedDict)

        return self._set_updated_data(data)

    def _set_updated_data(self, updated_data):
        data = OrderedDict()

        if '_meta' in updated_data.keys():
            updated_meta = []
            for row in updated_data['_meta']:
                if not row or not row[0]:
                    break

                if len(row) < 2:
                    updated_meta_value = None
                else:
                    updated_meta_value = row[1]

                self.meta[row[0]] = updated_meta_value

            updated_data.pop('_meta')

        for k, v in updated_data.items():
            data[k] = v

        return data

    @property
    def formatted_object(self):
        formatted_object = OrderedDict(
            _meta=list(self.meta.items())
        )

        for sheet_name, sheet_data in self.data.items():
            formatted_sheet_object = []
            for row in sheet_data:
                formatted_sheet_object.append(RowExport(row))
            formatted_object[sheet_name] = formatted_sheet_object

        return formatted_object

    def save(self, out_file: str, retain_meta=True, out_format=None):
        self.meta['modified'] = datetime.fromtimestamp(datetime.now().timestamp()).isoformat()
        self.meta.move_to_end('modified', last=False)

        if 'created' in self.meta.keys():
            self.meta.move_to_end('created', last=False)

        if out_format is None:
            out_base, out_format = os.path.splitext(out_file)
        else:
            out_base = os.path.splitext(out_file)[0]

        save_data = copy.deepcopy(self.data)
        if out_format == '.json' or retain_meta:
            save_data['_meta'] = list(self.meta.items())
            print(save_data)
            save_data.move_to_end('_meta', last=False)
        else:
            if '_meta' in save_data.keys():
                save_data.pop('_meta')

        for sheet_name, sheet_matrix in save_data.items():
            for i, row in enumerate(sheet_matrix):
                save_data[sheet_name][i] = RowExport(row)

        if out_format == '.ods':
            self._save_pyexcel_ods(out_file=out_file, out_data=save_data)
        elif out_format == '.xlsx':
            self._save_openpyxl(out_file=out_file, out_data=save_data)
        elif out_format == '.json':
            if os.path.splitext(out_base)[1] == '.pyexcel':
                self._save_pyexcel_json(out_file=out_file, out_data=save_data)
            else:
                self._save_json(out_file=out_file, out_data=save_data)
        else:
            raise ValueError('Unsupported file format, {}.'.format(out_file))

    @staticmethod
    def _save_pyexcel_ods(out_file: str, out_data: OrderedDict):
        pyexcel.save_as(out_data, out_file)

    def _save_openpyxl(self, out_file: str, out_data: OrderedDict):
        openpyxl_export(out_data=out_data, out_file=out_file, meta=self.meta)

    @staticmethod
    def _save_pyexcel_json(out_file: str, out_data: OrderedDict):
        with open(out_file, 'w') as f:
            export_string = json.dumps(out_data, cls=PyexcelExportEncoder,
                                       indent=2, ensure_ascii=False)
            f.write(export_string)

    @staticmethod
    def _save_json(out_file: str, out_data: OrderedDict):
        with open(out_file, 'w') as f:
            export_string = json.dumps(out_data, cls=RowExportEncoder,
                                       indent=2, ensure_ascii=False)
            f.write(export_string)
