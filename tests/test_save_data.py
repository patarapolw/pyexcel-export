import pytest
from pathlib import Path

import shutil
import pyexcel
import pyexcel_export
import logging

debugger_logger = logging.getLogger('debug')


@pytest.mark.parametrize('in_file', [
    'big_data.xlsx',
    'with_meta.xlsx'
])
@pytest.mark.parametrize('out_format', [
    '.pyexcel.json',
    '.yaml'
])
@pytest.mark.parametrize('retain_meta', [True, False])
def test_save_new_yaml_json(in_file, out_format, retain_meta, test_file, out_file, request):
    """

    :param str in_file:
    :param str out_format:
    :param bool retain_meta:
    :param test_file: function defined in conftest.py
    :param out_file: function defined in conftest.py
    :return:
    """
    in_file = test_file(in_file)

    assert isinstance(in_file, Path)

    data, meta = pyexcel_export.get_data(in_file)

    pyexcel_export.save_data(out_file(out_base=request.node.name, out_format=out_format), data=data, meta=meta, retain_meta=retain_meta)


@pytest.mark.parametrize('in_file', [
    'big_data.xlsx',
    'with_meta.xlsx'
])
@pytest.mark.parametrize('retain_meta', [True, False])
def test_save_new_excel(in_file, retain_meta, test_file, out_file, request):
    """

    :param str in_file:
    :param bool retain_meta:
    :param test_file: function defined in conftest.py
    :param out_file: function defined in conftest.py
    :return:
    """
    in_file = test_file(in_file)

    assert isinstance(in_file, Path)

    data, meta = pyexcel_export.get_data(in_file)

    pyexcel_export.save_data(out_file(out_base=request.node.name, out_format='.xlsx'), data, meta=meta, retain_meta=retain_meta)


# @pytest.mark.parametrize('in_file', [
#     'test.xlsx'
# ])
# @pytest.mark.parametrize('out_format', [
#     '.pyexcel.json',
#     '.yaml'
# ])
# def test_save_retain_styles(in_file, out_format, test_file, out_file):
#     in_file = test_file(in_file)
#
#     data, meta = pyexcel_export.get_data(in_file)
#
#     pyexcel_export.save_data(out_file(out_format=out_format), data,
#                              meta=meta, retain_meta=True, retain_styles=True)


@pytest.mark.parametrize('in_file', [
    'big_data.xlsx'
])
def test_update(in_file, test_file, out_file, request):
    in_file = test_file(in_file)

    assert isinstance(in_file, Path)

    shutil.copy(str(in_file.absolute()), str(out_file(out_base=request.node.name, out_format=".xlsx").absolute()))

    data, meta = pyexcel_export.get_data(in_file)

    data['TagDict'].append([
        'new rosai-dorfman disease',
        '''Typically massive lymph node involvement that may also involve extranodal sites
Associated with systemic symptoms (fever, leukocytosis, anemia)
Large histiocytes with abundant pale eosinophilic cytoplasm, mildly atypical round
vesicular nuclei
Lymphocytophagocytosis (emperipolesis) in background of mature lymphocytes and
plasma cells''',
        '“Dorfman disease” emperipolesis, HNB'
    ])

    pyexcel_export.save_data(out_file(out_base=request.node.name, out_format=".xlsx", do_overwrite=True), data, meta=meta)


@pytest.mark.parametrize('in_file', [
    'big_data.xlsx'
])
def test_update_classic(in_file, test_file, out_file, request):
    in_file = test_file(in_file)

    assert isinstance(in_file, Path)

    shutil.copy(str(in_file.absolute()), str(out_file(out_base=request.node.name, out_format=".xlsx").absolute()))

    data = pyexcel.get_book_dict(file_name=str(in_file.absolute()))
    debugger_logger.debug(data)

    data['TagDict'].append([
        'new rosai-dorfman disease',
        '''Typically massive lymph node involvement that may also involve extranodal sites
Associated with systemic symptoms (fever, leukocytosis, anemia)
Large histiocytes with abundant pale eosinophilic cytoplasm, mildly atypical round
vesicular nuclei
Lymphocytophagocytosis (emperipolesis) in background of mature lymphocytes and
plasma cells''',
        '“Dorfman disease” emperipolesis, HNB'
    ])

    pyexcel_export.save_data(out_file(out_base=request.node.name, out_format=".xlsx", do_overwrite=True), data)
