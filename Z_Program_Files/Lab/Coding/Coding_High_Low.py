# Excel file extraction

import xlrd
file_location = 'C:/GDS_Python/Lab/KOJ.xls'
workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_name('Sheet1')
KOJ_excel = worksheet._cell_values

gender ="f"

for row in (KOJ_excel[1:59]):
    keyword = row[0]
    trow = (row)
    del (trow[0:2])
    del (trow[1])
    del (trow[3:5])
    del (trow[5:])
    row = trow


# reference high vs low

        # srow = (row[4].split("-"))
        # print(srow)
        # value1 = (float(srow[1]) - float(row[2]))
        # value2 = (float(srow[0]) - float(row[2]))
        # # print(value1,value2)
        #
        # if (100 + value1 < 100) :
        #     row.append("High")
        #     print(row)
        # elif (100 + value2 > 100) :
        #     row.append("Low")
        #     print(row)
        # else:
        #     row.append(".")
        #     print(row)


