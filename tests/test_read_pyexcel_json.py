import pytest

from collections import OrderedDict

from pyexcel_export import read_pyexcel_json
from pyexcel_export.defaults import Meta

from tests import getfile


@pytest.mark.parametrize('in_file', [
    'template.pyexcel.json'
])
def test_read_pyexcel_json(in_file):
    data, meta = read_pyexcel_json(getfile(in_file))

    assert isinstance(data, OrderedDict)
    for k, v in data.items():
        assert isinstance(v, list)
        for row in v:
            assert isinstance(row, list)
            for cell in row:
                assert isinstance(cell, (int, str, bool, float))

    assert isinstance(meta, Meta)
