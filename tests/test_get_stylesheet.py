import pytest

import pyexcel_export
from pyexcel_export.defaults import Meta

from tests import getfile


@pytest.mark.parametrize('in_file', [
    None,
    'template.pyexcel.json',
    'template.xlsx',
    'update_to.xlsx'
])
def test_get_stylesheet(in_file):
    if in_file is not None:
        in_file = getfile(in_file)

    stylesheet = pyexcel_export.get_stylesheet(in_file)

    assert isinstance(stylesheet, Meta)
