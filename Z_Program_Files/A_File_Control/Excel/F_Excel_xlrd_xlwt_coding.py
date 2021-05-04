from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy

def file_save(x):
    with open('RC_result_file.txt', 'a', encoding='utf-8') as f:
        for item in (x):
            f.writelines(item)

wb = open_workbook("Labgen.xlsx",'a')
ws = wb.sheet_by_index(0)

for i in range(300):
    name = (ws.cell(i,2).value)
    if name == '최재천':
        print(ws.row_values(i)[1:])
        row = (ws.row_values(i)[1:])
        file_save({str(row)+'\n'})