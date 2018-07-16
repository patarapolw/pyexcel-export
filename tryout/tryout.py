from conftest import test_file
from pyexcel_export import ExcelLoader


if __name__ == '__main__':
   loader = ExcelLoader(test_file()("test.xlsx"))
   print(dict([(k, v) for k, v in loader.__dict__.items() if k != 'meta']))
   print(loader.meta)
