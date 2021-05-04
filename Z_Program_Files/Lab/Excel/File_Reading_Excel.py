# Excel file extraction

import xlrd
file_location = 'C:/GDS_Python/Excel/KOJ.xls'
workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_name('Sheet1')
KOJ_excel = worksheet._cell_values

for row in (KOJ_excel[1:6]):
    keyword = row[0]
    print(row)
    # print(type(row))
    with open('c:\\GDS_Python\\Excel\\result.txt', 'a', encoding='utf8') as outfile:
         outfile.writelines('%s\n' % row)
    outfile.close()


 # xlrd file reading and writing with space
# fhand = open('c:\\GDS_Python\\Excel\\result.txt','r', encoding='utf8')
# print("-"*40)
# for labres in fhand:
#     print(labres)
#     words = labres.split()
#     word = words[3:]
#     print(word)
# fhand.close()