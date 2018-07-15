import openpyxl


if __name__ == '__main__':
   wb = openpyxl.load_workbook('long.xlsx')
   print(wb['_meta']['B3'].__dict__)

   # wb.save('test.xlsx')
