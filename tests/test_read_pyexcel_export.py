import pytest
import os

from collections import OrderedDict

import pyexcel_export


@pytest.mark.parametrize('in_file', [
    'data_only.pyexcel.json',
    'no_style.pyexcel.json',
    'full.pyexcel.json',
])
def test_read_pyexcel_json(in_file, test_file, out_file):
    data, meta = pyexcel_export.get_data(test_file(in_file))

    in_base = os.path.split(os.path.splitext(in_file)[0])[1]

    assert isinstance(data, OrderedDict)
    for k, v in data.items():
        assert isinstance(v, list)
        for row in v:
            assert isinstance(row, list)
            for cell in row:
                assert isinstance(cell, (int, str, bool, float))

    pyexcel_export.save_data(out_file(in_base, ".xlsx"), data=data, meta=meta)


@pytest.mark.parametrize('in_file', [
    'data_only.yaml',
    'no_style.yaml',
    'full.yaml'
])
def test_read_pyexcel_yaml(in_file, test_file, out_file):
    data, meta = pyexcel_export.get_data(test_file(in_file))

    in_base = os.path.split(os.path.splitext(in_file)[0])[1]

    assert isinstance(data, OrderedDict)
    for k, v in data.items():
        assert isinstance(v, list)
        for row in v:
            assert isinstance(row, list)
            for cell in row:
                assert isinstance(cell, (int, str, bool, float))

    pyexcel_export.save_data(out_file(in_base, ".xlsx"), data=data, meta=meta)
