import pytest

import pyexcel_export
from pyexcel_export.defaults import Meta

from tests import getfile


@pytest.mark.parametrize('in_file', [
    None,
    'data_only.pyexcel.json',
    'data_only.yaml',
    'full.pyexcel.json',
    'full.yaml',
    'no_style.pyexcel.json',
    'no_style.yaml',
    'test.xlsx'
])
def test_get_meta(in_file):
    if in_file is not None:
        in_file = getfile(in_file)

    meta = pyexcel_export.get_meta(in_file)

    assert isinstance(meta, Meta)
