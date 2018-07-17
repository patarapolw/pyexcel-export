import pytest
from pathlib import Path

import pyexcel_export
from pyexcel_export.defaults import Meta


@pytest.mark.parametrize('in_file', [
    None,
    'data_only.pyexcel.json',
    'data_only.yaml',
    'full.pyexcel.json',
    'full.yaml',
    'no_style.pyexcel.json',
    'no_style.yaml',
    'big_data.xlsx',
    'with_meta.xlsx'
])
def test_get_meta(in_file, test_file):
    if in_file is not None:
        in_file = test_file(in_file)

    assert in_file is None or isinstance(in_file, Path)

    meta = pyexcel_export.get_meta(in_file)

    assert isinstance(meta, Meta)
