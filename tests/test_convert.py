import pytest

import os
import shutil

from pyexcel_export.app import ExcelLoader
from pyexcel_export.dir import root_path


@pytest.mark.parametrize('in_file, out_format', [
    ('template.xlsx', '.pyexcel.json'),
    ('template.pyexcel.json', '.xlsx')
])
def test_save_new(in_file, out_format):
    in_base = os.path.splitext(in_file)[0]
    if os.path.splitext(in_base)[1] == '.pyexcel':
        in_base = os.path.splitext(in_base)[0]

    while os.path.exists(root_path('tests/output/{}{}'.format(in_base, out_format))):
        out_format = '_' + out_format

    excel_loader = ExcelLoader(root_path('tests/input/{}'.format(in_file)))
    excel_loader.save(root_path('tests/output/{}{}'.format(in_base, out_format)))


@pytest.mark.parametrize('in_file, update_to', [
    ('template.pyexcel.json', 'update_to.xlsx')
])
def test_update(in_file, update_to):
    shutil.copy(root_path('tests/input/{}'.format(update_to)),
                root_path('tests/output/{}'.format('updated.xlsx')))

    excel_loader = ExcelLoader(root_path('tests/input/{}'.format(in_file)))
    excel_loader.save(root_path('tests/output/{}'.format('updated.xlsx')))
