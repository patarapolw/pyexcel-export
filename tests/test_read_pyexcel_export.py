import pytest
from pathlib import Path

from collections import OrderedDict
from pathlib import Path

import pyexcel_export


@pytest.mark.parametrize('in_file', [
    'test.pyexcel.json',
    'test.json'
])
def test_read_pyexcel_json(in_file, test_file, out_file, request):
    """

    :param str in_file:
    :param test_file: function defined in conftest.py
    :param out_file: function defined in conftest.py
    :return:
    """
    in_file = test_file(in_file)

    assert isinstance(in_file, Path)

    data, meta = pyexcel_export.get_data(in_file)

    assert isinstance(data, OrderedDict)
    for k, v in data.items():
        assert isinstance(v, list)
        for row in v:
            assert isinstance(row, list)
            for cell in row:
                assert isinstance(cell, (int, str, bool, float))

    pyexcel_export.save_data(out_file(out_base=request.node.name, out_format='.xlsx'), data=data, meta=meta)


@pytest.mark.parametrize('in_file', [
    'test.yaml'
])
def test_read_pyexcel_yaml(in_file, test_file, out_file, request):
    """

    :param str in_file:
    :param test_file: function defined in conftest.py
    :param out_file: function defined in conftest.py
    :return:
    """
    in_file = test_file(in_file)

    assert isinstance(in_file, Path)

    data, meta = pyexcel_export.get_data(in_file)

    assert isinstance(data, OrderedDict)
    for k, v in data.items():
        assert isinstance(v, list)
        for row in v:
            assert isinstance(row, list)
            for cell in row:
                assert isinstance(cell, (int, str, bool, float))

    pyexcel_export.save_data(out_file(out_base=request.node.name, out_format=".xlsx"), data=data, meta=meta)
