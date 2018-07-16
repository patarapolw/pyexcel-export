import pytest
from pathlib import Path

from pyexcel_export import ExcelLoader


@pytest.mark.parametrize('in_file', [
    "data_only.pyexcel.json",
    "data_only.yaml",
    "full.pyexcel.json",
    "full.yaml",
    "no_style.pyexcel.json",
    "no_style.yaml",
    "test.xlsx"
])
@pytest.mark.parametrize('out_format', [
    ".xlsx",
    ".json",
    ".pyexcel.json",
    ".yaml"
])
def test_excel_loader(in_file, out_format, test_file, out_file):
    """

    :param str in_file:
    :param str out_format:
    :param test_file: function defined in conftest.py
    :param out_file: function defined in conftest.py
    :return:
    """
    in_file = test_file(in_file)
    assert isinstance(in_file, Path)

    loader = ExcelLoader(in_file=in_file)
    loader.save(out_file=out_file(in_file.stem + '_excel_loader', '.xlsx'))
