from .app import ExcelLoader
from .defaults import Meta


def get_data(in_file, **flags):
    loader = ExcelLoader(in_file, **flags)

    return loader.data, loader.meta


def get_stylesheet(in_file=None, **flags):
    """

    :param in_file: supported file formats are *.xlsx, *.json and *.pyexcel.json
    :param flags:
    :return:
    """
    if in_file:
        loader = ExcelLoader(in_file, **flags)

        return loader.meta
    else:
        return Meta(**flags)


def save_data(out_file, data, stylesheet=None, retain_styles=False, **flags):
    """

    :param out_file: supported file formats are *.xlsx, *.json and *.pyexcel.json
    :param data:
    :param stylesheet:
    :param retain_styles: whether you want to retain the overwritten worksheet's formatting
    :param flags
    :return:
    """
    loader = ExcelLoader(**flags)
    loader.meta = stylesheet
    loader.data = data
    loader.save(out_file, retain_styles=retain_styles)
