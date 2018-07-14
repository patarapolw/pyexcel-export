import pytest

from pyexcel_export.app import ExcelLoader
from pyexcel_export.formatter import ExcelFormatter
from pyexcel_export.dir import root_path


@pytest.mark.parametrize('in_file', [
    # 'template.pyexcel.json',
    'template.xlsx'
])
def test_data(in_file):
    excel_loader = ExcelLoader(root_path('tests/input/{}'.format(in_file)))
    # print(excel_loader.data)
    print(excel_loader.meta.view)


@pytest.mark.parametrize('in_file', [
    'template.xlsx'
])
def test_formatter(in_file):
    print(type(ExcelFormatter(root_path('tests/input/{}'.format(in_file))).data))
