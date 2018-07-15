import pyexcel_export
from tests import getfile


if __name__ == '__main__':
    data, meta = pyexcel_export.get_data(getfile('test.xlsx'))
    print(meta)

    pyexcel_export.save_data('test.yaml', data, meta, retain_styles=True)
