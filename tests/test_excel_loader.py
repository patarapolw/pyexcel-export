import pytest

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
    loader = ExcelLoader(in_file=test_file(in_file))
    loader.save(out_file=out_file("test", out_format))
