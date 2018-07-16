from .app import ExcelLoader
from .defaults import Meta


def get_data(in_file, **flags):
    """

    :param str|Path in_file: supported file formats are *.xlsx, *.yaml, *.json and *.pyexcel.json
    :param flags: currently supported flags are 'has_header', 'freeze_header', 'col_width_fit_param_keys',
                  'col_width_fit_ids', 'bool_as_string', 'allow_table_hiding'
    :return:
    """
    loader = ExcelLoader(in_file, **flags)

    return loader.data, loader.meta


def get_meta(in_file=None, **flags):
    """

    :param str|Path in_file:
    :param flags:
    :return:
    """
    if in_file:
        loader = ExcelLoader(in_file, **flags)

        return loader.meta
    else:
        return Meta(**flags)


def save_data(out_file, data, meta=None, retain_meta=False, retain_styles=False, **flags):
    """

    :param str|Path out_file: supported file formats are *.xlsx, *.json and *.pyexcel.json
    :param data:
    :param meta:
    :param retain_styles: whether you want to retain the overwritten worksheet's formatting
    :param retain_meta: whether you want to write _meta to your file
    :param flags:
    :return:
    """
    loader = ExcelLoader(**flags)

    if meta is not None:
        loader.meta = meta

    loader.data = data
    loader.save(out_file, retain_styles=retain_styles, retain_meta=retain_meta)
