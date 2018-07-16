from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

if __name__ == '__main__':
    p = Path(__file__).parent.joinpath(Path('test.xlsx'))
    wb = openpyxl.load_workbook(p)
    print(wb.sheetnames)
    ws = wb['test']
    for i in range(len(next(ws.iter_cols()))):
        print(i, ws.column_dimensions[get_column_letter(i+1)].width)

    for j, row in enumerate(ws):
        for i, cell in enumerate(row):
            cell.alignment = Alignment(wrap_text=True)
            if cell.value is not None:
                print(i, j, len(str(cell.value)))

        print(ws.row_dimensions[j + 1].height)
