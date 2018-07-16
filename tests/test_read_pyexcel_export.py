import pytest
import os

from collections import OrderedDict

import pyexcel_export

from tests import getfile


@pytest.mark.parametrize('in_file', [
    'data_only.pyexcel.json',
    'data_only.yaml',
    'no_style.pyexcel.json',
    'no_style.yaml',
    'full.pyexcel.json',
    'full.yaml'
])
def test_read_pyexcel_export(in_file, tmpdir):
    data, meta = pyexcel_export.get_data(getfile(in_file))

    in_base = os.path.split(os.path.splitext(in_file)[0])[1]

    assert isinstance(data, OrderedDict)
    for k, v in data.items():
        assert isinstance(v, list)
        for row in v:
            assert isinstance(row, list)
            for cell in row:
                assert isinstance(cell, (int, str, bool, float))

    pyexcel_export.save_data(tmpdir.join(in_base + ".xlsx"), data=data, meta=meta)
