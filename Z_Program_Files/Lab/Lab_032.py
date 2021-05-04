# Excel file extraction

import xlrd
file_location = 'C:/GDS_Python/Lab/KOJ.xls'
workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_name('Sheet1')
KOJ_excel = worksheet._cell_values

print("-"*75)
# Excel selection items
gender = input ("Gender 성별 (m/f)  ?   ")
print("-"*75)

for row in (KOJ_excel[1:64]):
    keyword = row[0]
    trow = (row)
    del (trow[0:2])
    del (trow[1])
    del (trow[3:5])
    del (trow[5:])
    trow[1].strip()
    row = trow

    def charact_lab(x):
        if row[1].startswith(x):
            row.insert(0,"AAA")
            crow = row.copy()
            # print(crow)
            return crow
    charact_lab('ABO')
    charact_lab('Rh')
    charact_lab('RF')
    charact_lab('RPR')
    charact_lab('HIV')
    charact_lab('HAV')
    charact_lab('HCV')
    charact_lab('HBs')
    charact_lab('Diff')

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
    change_items('ALP','35-130')
    change_items('Sodium','135-145')
    change_items('Chloride','98-110')
    change_items('WBC','4.00-11.00')
    change_items('MCV','80-100')
    change_items('MCH','26.0-34.0')
    change_items('Other','0-1')

    change_items('RBC', '4.5-5.9')
    change_items('Hemoglobin', '12.5-17.5')
    change_items('Hematocrit', '38.0-53.0')
    # change_items('Potassium',)


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
#       print(row)


# result printing
# reference high vs low
    if row[0] != "AAA":
        srow = (row[4].split("-"))

        value1 = (float(srow[1]) - float(row[2]))
        value2 = (float(srow[0]) - float(row[2]))

        if (100 + value1 < 100):
            row.append("High")
            # print(row)
        elif (100 + value2 > 100):
            row.append("Low")
            # print(row)
        else:
            row.append(".")
            # print(row)

# printing
    if row[0] != "AAA":
        print(row[0].ljust(7), row[1].ljust(25), row[2].rjust(12), "  ", row[3].ljust(15), row[4].rjust(10),"   ", row[5].ljust(10))

print("-"*75)
# character laboratory data

for row in (KOJ_excel[1:64]):
    keyword = row[0]
    trow = (row)
    # del (trow[0:2])
    # del (trow[1])
    # del (trow[3:5])
    # del (trow[5:])
    trow[1].rstrip()
    row = trow
    if row[0] == 'AAA':
        del row[0]
        print(row[0].ljust(5), row[1].ljust(21), row[2].rjust(18), "  ", row[3].ljust(10),row[4].ljust(18))

print("-"*75)
# urine laboratory data

for row in (KOJ_excel[64:87]):
    keyword = row[0]
    trow = (row)
    del (trow[0:2])
    del (trow[1])
    del (trow[3:5])
    del (trow[5:])
    trow[1].rstrip()
    row = trow
    print(row[0].ljust(7), row[1].ljust(25), row[2].rjust(12), "  ", row[3].ljust(15), row[4].rjust(10), "   ",)
print("-"*75)