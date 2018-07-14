import pytest

from pyexcel_formatter.app import ExcelFormatter
from pyexcel_formatter.dir import root_path


@pytest.mark.parametrize('in_file', [
    'template.pyexcel.json'
])
def test_data(in_file):
    formatter = ExcelFormatter(root_path('tests/input/{}'.format(in_file)))
    print(formatter.data)
    print(formatter.meta)
