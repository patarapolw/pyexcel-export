import pytest

import os
import shutil
import json

import pyexcel
import pyexcel_export

from tests import getfile


@pytest.mark.parametrize('in_file', ['test.xlsx'])
@pytest.mark.parametrize('out_format', [
    '.pyexcel.json',
    '.yaml',
    '.xlsx'
])
@pytest.mark.parametrize('retain_meta', [True, False])
def test_save_new(in_file, out_format, retain_meta):
    data, meta = pyexcel_export.get_data(getfile(in_file, 'input'))

    out_base = os.path.splitext(in_file)[0]
    out_file = out_base + out_format

    while os.path.exists(getfile(out_file, 'output')):
        out_file = '_' + out_file

    pyexcel_export.save_data(getfile(out_file, 'output'), data, meta=meta, retain_meta=retain_meta)


@pytest.mark.parametrize('in_file', ['test.xlsx'])
@pytest.mark.parametrize('out_format', [
    '.pyexcel.json',
    '.yaml'
])
def test_save_retain_styles(in_file, out_format):
    data, meta = pyexcel_export.get_data(getfile(in_file, 'input'))

    out_base = os.path.splitext(in_file)[0]
    out_file = out_base + out_format

    while os.path.exists(getfile(out_file, 'output')):
        out_file = '_' + out_file

    pyexcel_export.save_data(getfile(out_file, 'output'), data, meta=meta, retain_meta=True, retain_styles=True)


@pytest.mark.parametrize('in_file, update_to', [
    ('test.xlsx', 'update_to.xlsx')
])
def test_update(in_file, update_to):
    shutil.copy(getfile(update_to, 'input'), getfile(update_to, 'output'))

    data, meta = pyexcel_export.get_data(getfile(in_file, 'input'))

    pyexcel_export.save_data(getfile(update_to, 'output'), data, meta=meta)
