# Excel file extraction

import xlrd
file_location = 'C:/GDS_Python/Lab/KOJ.xls'
workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_name('Sheet1')
KOJ_excel = worksheet._cell_values

for row in (KOJ_excel[1:88]):
    keyword = row[0]
    print(row[5])
    # print (list(enumerate(row))
    # print(type(row))
    with open('c:\\GDS_Python\\Lab\\result.txt', 'a', encoding='utf8') as outfile:
        outfile.writelines('%s\n' % ("{:<15}".format(row[4]) +"{:>15}".format(row[5]) +"{:>7}".format(row[8]) +"{:>20}".format(row[9])))
        # print("{:<15}".format(row[4]) +"{:>15}".format(row[5]) +"{:>7}".format(row[8]) +"{:>20}".format(row[9]))
    outfile.close()
#
# print("-"*30)
# for row in (KOJ_excel[1:11]):
#     keyword = row[0]
#
#     print(refdat[0])
#     refdat = ['6.4-8.2', '3.8-5.3', '0.20-1.20', '8.0-23.0', '0.30-1.30', '3.9-8.3', '0-40', '0-40', '35-130', '0-73']
#     highlow  = refdat.split("-")
#     print(row[5])