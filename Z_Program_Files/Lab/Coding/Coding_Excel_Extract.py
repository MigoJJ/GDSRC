# Excel file extraction

import xlrd

file_location = 'C:/GDS_Python/Lab/KOJ.xls'
workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_name('Sheet1')
KOJ_excel = worksheet._cell_values

gender = "f"

for row in (KOJ_excel[1:64]):
    keyword = row[0]
    trow = (row)
    del (trow[0:2])
    del (trow[1])
    del (trow[3:5])
    del (trow[5:])
    row = trow
    # print(row)


# ABO  Rh  list control

    # for i in ['ABO','Rh ','RF ','RPR','HIV','HAV','HBs','HCV']:
    #     if row[1].startswith(i):
