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

# common reference replace

# data change reference <   M:   >

    def change_ref (x):
        if row[4].startswith(x):
            if row[4].startswith('M:'):
                row.insert(4, row[4][2:])  # List 4 item  from 2 character
            if row[4].startswith('성인'):
                row.insert(4, row[4][3:])  # List 4 item  from 2 character
            row.pop()
            # print(row)
            return row

    change_ref('M:')
    change_ref('성인')


# common reference replace   <   \n   >
#     def remove_n(x):
#         J = row.pop()
#         # print(J)
#         row.append(J)
#
#
#     remove_n(row[4])


# <  common  > change items
    def change_items (x,y):
        if row[1].startswith(x):
            row.insert(4,y)
            del row[5]
            return row

    change_items('BUN','8.0-23.0')
    change_items('Amy','28-110')
    change_items('Chol','110-200')
    change_items('Tri','0-150')
    change_items('HDL','0-40')
    change_items('LDL','0-100')
    change_items('Gluc','70-100')
    change_items('25-OH','20-100')
    change_items('AFP','0-8.1')
    change_items('CEA','0-5.0')
    change_items('CA19-9','0-37.0')
    change_items('PSA','0-4')
    change_items('ESR','0-15')
    change_items('GGT', '0-73')
    # ...
    # ...
    # print(row)
    print("-"*60)

# <  female  > change items
    if gender =="f":
        def change_f(x, y):
            if row[1].startswith(x):
                row.insert(4, y)
                del row[5]
                return row
        change_f('Uric', '2.8-6.1')
        change_f('GGT', '0-38')
        change_f('CK', '34-145')
        change_f('Iron', '32-153')
        change_f('RBC', '3.70-5.2')
        change_f('Hemoglobin', '11.0-16.0')
        change_f('Hematocrit', '36.0-46.0')
        change_f('ESR', '0-20')
        # print(row)
        

    # corrected data presentation

# reference high vs low

        srow = (row[4].split("-"))
        # print(srow)
        value1 = (float(srow[1]) - float(row[2]))
        value2 = (float(srow[0]) - float(row[2]))
        # print(value1,value2)

        if (100 + value1 < 100) :
            row.append("High")
            print(row)
        elif (100 + value2 > 100) :
            row.append("Low")
            print(row)
        else:
            row.append(".")
            print(row)

        print(row[0].ljust(12), row[1].ljust(18), row[2].rjust(16), "  ", row[3].ljust(10),row[4].rjust(15),row[5].rjust(10))
