import pytest

import os
import shutil
import pyexcel_export


@pytest.mark.parametrize('in_file', ['test.xlsx'])
@pytest.mark.parametrize('out_format', [
    '.pyexcel.json',
    '.yaml',
    '.xlsx'
])
@pytest.mark.parametrize('retain_meta', [True, False])
def test_save_new(in_file, out_format, retain_meta, test_file, out_file):
    data, meta = pyexcel_export.get_data(test_file(in_file))

    out_base = os.path.splitext(in_file)[0]
    out_filename = out_base + out_format

    pyexcel_export.save_data(out_file(out_filename), data, meta=meta, retain_meta=retain_meta)


@pytest.mark.parametrize('in_file', ['test.xlsx'])
@pytest.mark.parametrize('out_format', [
    '.pyexcel.json',
    '.yaml'
])
def test_save_retain_styles(in_file, out_format, test_file, out_file):
    data, meta = pyexcel_export.get_data(test_file(in_file))

    out_base = os.path.splitext(in_file)[0]
    out_filename = out_base + out_format

    pyexcel_export.save_data(out_file(out_filename), data, meta=meta, retain_meta=True, retain_styles=True)


@pytest.mark.parametrize('in_file, update_to', [
    ('test.xlsx', 'update_to.xlsx')
])
def test_update(in_file, update_to, test_file, out_file):
    shutil.copy(test_file(update_to), out_file(update_to))

    data, meta = pyexcel_export.get_data(test_file(in_file))

    pyexcel_export.save_data(out_file(update_to, do_overwrite=True), data, meta=meta)
