import pyexcel
from collections import OrderedDict
import json
import copy
import oyaml as yaml
import logging
from pathlib import Path
import re

from .serialize import RowExport, PyexcelExportEncoder, MyEncoder
from .yaml_serialize import PyExcelYamlLoader
from .defaults import Meta
from .formatter import ExcelFormatter

debugger_logger = logging.getLogger('debug')


class ExcelLoader:
    def __init__(self, in_file=None, **flags):
        """

        :param str|Path|None in_file:
        :param flags:
        """
        if in_file:
            if not isinstance(in_file, Path):
                in_file = Path(in_file)

            self.in_file = in_file
            self.meta = Meta(**flags)

            in_format = in_file.suffixes

            if in_format and in_format[-1] == '.xlsx':
                self.data = self._load_pyexcel_xlsx()
            elif in_format[-1] == '.json':
                if len(in_format) > 1 and in_format[-2] == '.pyexcel':
                    self.data = self._load_pyexcel_json()
                else:
                    self.data = self._load_json()
            elif in_format[-1] in ('.yaml', '.yml'):
                self.data = self._load_yaml()
            else:
                raise ValueError('Unsupported file format, {}.'.format(in_format))
        else:
            self.meta = Meta(**flags)

    def _load_pyexcel_xlsx(self):
        updated_data = pyexcel.get_book_dict(file_name=str(self.in_file.absolute()))
        self.meta.setdefault('_styles', dict())['excel'] = ExcelFormatter(self.in_file).data

        return self._set_updated_data(updated_data)

    def _load_pyexcel_json(self):
        with self.in_file.open(encoding='utf-8') as f:
            data = json.load(f, object_pairs_hook=OrderedDict)

        for sheet_name, sheet_data in data.items():
            for j, row in enumerate(sheet_data):
                for i, cell in enumerate(row):
                    data[sheet_name][j][i] = json.loads(cell)

        return self._set_updated_data(data)

    def _load_json(self):
        with self.in_file.open(encoding='utf8') as f:
            data = json.load(f, object_pairs_hook=OrderedDict)

        return self._set_updated_data(data)

    def _load_yaml(self):
        with self.in_file.open(encoding='utf8') as f:
            data = yaml.load(f, Loader=PyExcelYamlLoader)

        return self._set_updated_data(data)

    def _set_updated_data(self, updated_data):
        data = OrderedDict()

        if '_meta' in updated_data.keys():
            for row in updated_data['_meta']:
                if not row or not row[0]:
                    break

                if len(row) < 2:
                    updated_meta_value = None
                else:
                    try:
                        updated_meta_value = list(json.loads(row[1]).values())[0]
                    except (json.decoder.JSONDecodeError, TypeError, AttributeError):
                        updated_meta_value = row[1]

                self.meta[row[0]] = updated_meta_value

            updated_data.pop('_meta')

        for k, v in updated_data.items():
            data[k] = v

        return data

    def save(self, out_file: str, retain_meta=True, out_format=None, retain_styles=True):
        """

        :param str|Path out_file:
        :param retain_meta:
        :param out_format:
        :param retain_styles:
        :return:
        """
        if not isinstance(out_file, Path):
            out_file = Path(out_file)

        if out_format is None:
            out_format = out_file.suffixes
        else:
            out_format = re.findall('\.[^.]+', out_format)

        save_data = copy.deepcopy(self.data)

        if out_format[-1] == '.xlsx':
            if '_meta' in save_data.keys():
                save_data.pop('_meta')

            self._save_openpyxl(out_file=out_file, out_data=save_data, retain_meta=retain_meta)
        else:
            if retain_meta:
                save_data['_meta'] = self.meta.matrix
                if not retain_styles:
                    for i, row in enumerate(save_data['_meta']):
                        if row[0] == '_styles':
                            save_data['_meta'].pop(i)
                            break

                save_data.move_to_end('_meta', last=False)
            else:
                if '_meta' in save_data.keys():
                    save_data.pop('_meta')

            if out_format[-1] == '.json':
                if len(out_format) > 1 and out_format[-2] == '.pyexcel':
                    self._save_pyexcel_json(out_file=out_file, out_data=save_data)
                else:
                    self._save_json(out_file=out_file, out_data=save_data)
            elif out_format[-1] in ('.yaml', '.yml'):
                self._save_yaml(out_file=out_file, out_data=save_data)
            else:
                raise ValueError('Unsupported file format, {}.'.format(out_file))

    def _save_openpyxl(self, out_file, out_data: OrderedDict, retain_meta: bool=True):
        """

        :param str|Path out_file:
        :param out_data:
        :param retain_meta:
        :return:
        """
        if not isinstance(out_file, Path):
            out_file = Path(out_file)

        formatter = ExcelFormatter(out_file)
        if out_file.exists():
            self.meta.setdefault('_styles', dict())['excel'] = formatter.data

        formatter.save(out_data, out_file, meta=self.meta, retain_meta=retain_meta)

    @staticmethod
    def _save_pyexcel_json(out_file, out_data: OrderedDict):
        """

        :param str|Path out_file:
        :param out_data:
        :return:
        """
        for k, v in out_data.items():
            for i, row in enumerate(v):
                out_data[k][i] = RowExport(row)

        if not isinstance(out_file, Path):
            out_file = Path(out_file)

        with out_file.open('w', encoding='utf8') as f:
            export_string = json.dumps(out_data, cls=PyexcelExportEncoder,
                                       indent=2, ensure_ascii=False)
            f.write(export_string)

    @staticmethod
    def _save_json(out_file, out_data: OrderedDict):
        """

        :param str|Path out_file:
        :param out_data:
        :return:
        """
        if not isinstance(out_file, Path):
            out_file = Path(out_file)

        with out_file.open('w', encoding='utf8') as f:
            export_string = json.dumps(out_data, cls=MyEncoder,
                                       indent=2, ensure_ascii=False)
            f.write(export_string)

    @staticmethod
    def _save_yaml(out_file, out_data: OrderedDict):
        """

        :param str|Path out_file:
        :param out_data:
        :return:
        """
        if not isinstance(out_file, Path):
            out_file = Path(out_file)

        with out_file.open('w', encoding='utf8') as f:
            yaml.dump(out_data, f, allow_unicode=True)
