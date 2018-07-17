from pathlib import Path

import pyexcel_export

if __name__ == '__main__':
    data, meta = pyexcel_export.get_data("PathoDict.yaml")
    pyexcel_export.save_data('PathoDict.xlsx', data)
