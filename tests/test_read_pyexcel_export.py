import pytest

from collections import OrderedDict

import pyexcel_export
from pyexcel_export.defaults import Meta

from tests import getfile


@pytest.mark.parametrize('in_file', [
    'short.pyexcel.json',
    'short.yaml',
    'long.yaml'
])
def test_read_pyexcel_export(in_file):
    data, meta = pyexcel_export.get_data(getfile(in_file))

    assert isinstance(data, OrderedDict)
    for k, v in data.items():
        assert isinstance(v, list)
        for row in v:
            assert isinstance(row, list)
            for cell in row:
                assert isinstance(cell, (int, str, bool, float))

    assert isinstance(meta, Meta)
