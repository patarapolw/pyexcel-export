import pytest

import os
import shutil
import json

import pyexcel
import pyexcel_export

from tests import getfile


@pytest.mark.parametrize('in_file', ['template.pyexcel.json'])
@pytest.mark.parametrize('out_file', [
    'template.pyexcel.json',
    'template.xlsx'
])
def test_save_new(in_file, out_file):
    data, meta = pyexcel_export.read_pyexcel_json(getfile(in_file, 'input'))

    while os.path.exists(getfile(out_file, 'output')):
        out_file = '_' + out_file

    pyexcel_export.save_data(getfile(out_file, 'output'), data, stylesheet=meta)


@pytest.mark.parametrize('in_file, update_to', [
    ('template.pyexcel.json', 'update_to.xlsx')
])
def test_update(in_file, update_to):
    shutil.copy(getfile(update_to, 'input'), getfile(update_to, 'output'))

    data, meta = pyexcel_export.read_pyexcel_json(getfile(in_file, 'input'))

    pyexcel_export.save_data(getfile(update_to, 'output'), data, stylesheet=meta)
