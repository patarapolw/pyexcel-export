import pytest
from pathlib import Path


@pytest.fixture(scope="module")
def test_file():
    def _test_file(filename):
        """

        :param str|Path filename:
        :return:
        """
        return Path(__file__).parent.joinpath('tests/input/').joinpath(filename)

    return _test_file


@pytest.fixture(scope="module")
def out_file():
    def _out_file(out_base, out_format=None, do_overwrite=False):
        """

        :param Path|str out_base:
        :param out_format:
        :param do_overwrite:
        :return:
        """
        if not isinstance(out_base, Path):
            out_base = Path(out_base)

        if out_format is None:
            out_format = out_base.suffix

        out_filename = Path(__file__).parent.joinpath('tests/output/').joinpath(out_base.name).with_suffix(out_format)
        if not do_overwrite:
            while out_filename.exists():
                out_filename = out_filename.with_name('_' + out_filename.name)
        else:
            # assert os.path.exists(out_filename)
            pass

        return out_filename

    return _out_file
