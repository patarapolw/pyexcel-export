import pytest
from pathlib import Path

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
    """

    :param str in_file:
    :param str out_format:
    :param bool retain_meta:
    :param test_file: function defined in conftest.py
    :param out_file: function defined in conftest.py
    :return:
    """
    in_file = test_file(in_file)

    assert isinstance(in_file, Path)

    data, meta = pyexcel_export.get_data(in_file)

    pyexcel_export.save_data(out_file(in_file, out_format), data, meta=meta, retain_meta=retain_meta)


@pytest.mark.parametrize('in_file', ['test.xlsx'])
@pytest.mark.parametrize('out_format', [
    '.pyexcel.json',
    '.yaml'
])
def test_save_retain_styles(in_file, out_format, test_file, out_file):
    in_file = test_file(in_file)

    data, meta = pyexcel_export.get_data(in_file)

    pyexcel_export.save_data(out_file(in_file.stem + '_retain_styles', out_format), data, meta=meta, retain_meta=True, retain_styles=True)


@pytest.mark.parametrize('in_file, update_to', [
    ('test.xlsx', 'update_to.xlsx')
])
def test_update(in_file, update_to, test_file, out_file):
    update_to = test_file(update_to)

    assert isinstance(update_to, Path)

    shutil.copy(str(update_to.absolute()), str(out_file(update_to).absolute()))

    in_file = test_file(in_file)

    data, meta = pyexcel_export.get_data(in_file)

    pyexcel_export.save_data(out_file("updated", ".xlsx", do_overwrite=True), data, meta=meta)
