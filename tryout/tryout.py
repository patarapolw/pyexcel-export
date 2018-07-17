import openpyxl

from conftest import test_file


if __name__ == '__main__':
    wb = openpyxl.load_workbook(str(test_file()('with_meta.xlsx')))
    print(wb['good sheet'])

