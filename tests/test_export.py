import pytest

from pyexcel_formatter.app import ExcelFormatter
from pyexcel_formatter.dir import root_path


@pytest.mark.parametrize('in_file, out_file', [
    ('template.xlsx', 'template.pyexcel.json')
])
def test_save(in_file, out_file):
    formatter = ExcelFormatter(root_path('tests/input/{}'.format(in_file)))
    formatter.save(root_path('tests/output/{}'.format(out_file)))
