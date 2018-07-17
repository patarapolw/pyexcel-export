import pytest
from pathlib import Path
import logging

debugger_logger = logging.getLogger('debug')


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
def out_file(request):
    def _out_file(out_base=None, out_format=None, do_overwrite=False):
        """

        :param Path|str|None out_base:
        :param out_format:
        :param do_overwrite:
        :return:
        """
        if out_base is not None:
            out_base = Path(out_base)
        else:
            out_base = Path(request.node.name)

        if out_format is None:
            out_format = out_base.suffix

        out_filename = Path(__file__).parent.joinpath('tests/output/').joinpath(out_base.name).with_suffix(out_format)
        if not do_overwrite:
            while out_filename.exists():
                out_filename = out_filename.with_name(out_filename.stem + '_' + ''.join(out_filename.suffixes))
        else:
            # assert os.path.exists(out_filename)
            pass

        return out_filename

    return _out_file
