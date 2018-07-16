import pytest
import os

from pyexcel_export.dir import TRUE_ROOT


@pytest.fixture(scope="module")
def test_file():
    def _test_file(filename):
        return os.path.join(TRUE_ROOT, 'tests', 'input', filename)

    return _test_file


@pytest.fixture(scope="module")
def out_file():
    def _out_file(out_base, out_format, do_overwrite=False):
        filename = out_base + out_format

        out_filename = os.path.join(TRUE_ROOT, 'tests', 'output', filename)
        if not do_overwrite:
            while os.path.exists(out_filename):
                filename = '_' + filename
                out_filename = os.path.join(TRUE_ROOT, 'tests', 'output', filename)
        else:
            # assert os.path.exists(out_filename)
            pass

        return out_filename

    return _out_file
