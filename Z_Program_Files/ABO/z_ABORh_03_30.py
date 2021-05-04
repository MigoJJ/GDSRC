# https://medium.com/@paulrohan/python-list-vs-tuple-vs-dictionary-4a48655c7934

# Excel file extraction

import xlrd
file_location = 'KOJ.xls'
workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_name('Sheet1')
KOJ_excel = worksheet._cell_values

def file_save(x):
    with open('C:/GDSRC/RC_result_file.txt', 'a',encoding='utf-8') as f:
        for item in (x):
            f.writelines(item)

for row in (KOJ_excel[1:64]):
    keyword = row[0]

# --- << INSERT lAB DATA REFERENECE CHANGE >> ---

    lab_file_listing = {}
    lab_file_listing = dict({'code': row[4], 'value': row[5], 'unit': row[8], 'reference': row[9]})
    list_lab = list(lab_file_listing.values())
    # # file_save(list_lab)
    # print(list_lab)
# -------------------------------------------------ABO  Rh
    if list_lab[0] == ('ABO혈액형(자동화)'):
        ABO1 = list_lab[1]
    elif list_lab[0] == ('Rh type(자동화)'):
        Rh1 = list_lab[1]
        print('\n------------------------------------------------------ABO  Rh type\n')
        file_save('\n------------------------------------------------------ABO  Rh type\n\n')
        print('당신의 혈액 형은  {1} 형,  Rh {0} 입니다.\n'
              '''응급실에서 수혈이 필요한 경우 필요한 혈액형을 신속하게 찾아낼 수 도 있기 때문에 중요한 정보 입니다.\n'''.format(Rh1, ABO1))
        file_save('당신의 혈액 형은  {1} 형,  Rh {0} 입니다.\n'
              '''응급실에서 수혈이 필요한 경우 필요한 혈액형을 신속하게 찾아낼 수 도 있기 때문에 중요한 정보 입니다.\n'''.format(Rh1, ABO1))

        print('------------------------------------------------------Hepatitis\n')
        file_save('\n------------------------------------------------------Hepatitis\n\n')
    elif list_lab[0] == ('HAV-Ab IgG     '):
        HAV_Ab_IgG1 = list_lab[1]
        if HAV_Ab_IgG1.startswith('Pos'):
            print('A 간염의 항체가 있습니다. 예방 접종이 필요하지 않습니다.\n')
            file_save('A 간염의 항체가 있습니다. 예방 접종이 필요하지 않습니다.\n')
        else:
            print('A 간염의 항체가 없습니다. 예방 접종이 필요합니다.\n'
                  '''A 형간염 예방 접종 설명...\n''')
            file_save('A 간염의 항체가 없습니다. 예방 접종이 필요합니다.\n'
                  '''A 형간염 예방 접종 설명...\n''')

    elif list_lab[0] == ('HBsAg   '):
        HBsAg1 = list_lab[1]
        if HBsAg1.startswith('Pos'):
            print('B 간염의 항원(Ag ,Antigen)이 있습니다.  정기적인 전문의의 진료를 받도록하십오.\n')
            file_save('B 간염의 항원(Ag ,Antigen)이 있습니다.  정기적인 전문의의 진료를 받도록하십오.\n')
    elif list_lab[0] == ('HBsAb  '):
        HBsAb1 = list_lab[1]
        if HBsAb1.startswith('Pos'):
            print('B 간염의 항체(Ab ,Antibody)가 있습니다.  B형 간염에 대한 예방접종이 필요 없습니다.\n')
            file_save('B 간염의 항체(Ab ,Antibody)가 있습니다.  B형 간염에 대한 예방접종이 필요 없습니다.\n')
        else:
            print('B 간염의 항체가 없습니다. 예방 접종이 필요합니다.\n'
                  '''B 형간염 예방 접종 설명...\n''')
            file_save('B 간염의 항체가 없습니다. 예방 접종이 필요합니다.\n'
                  '''B 형간염 예방 접종 설명...\n''')
    
    elif list_lab[0] == ('HCV Ab(일반)'):
        HCV1 = list_lab[1]
        if HCV1.startswith('Neg'):
            print('C 형 간염이 없습니다.\n')
            file_save('C 형 간염이 없습니다.\n')
        else:
            print('C 간염의 항체(Ab ,Antibody)가 있습니다. \n  C 형 간염의 확정적인 검사는 아닙니다. 전문의의 진료를 받도록하십오.\n')
            file_save('C 간염의 항체(Ab ,Antibody)가 있습니다. \n  C 형 간염의 확정적인 검사는 아닙니다. 전문의의 진료를 받도록하십오.\n')

    # ----------------dictionary definition and printing
    # lab_file_listing = {}
    # lab_file_listing = dict({'code': row[4], 'value': row[5], 'unit': row[8], 'reference': row[9]})
    # print(lab_file_listing)

    # print(lab_file_listing.keys())
    # print(lab_file_listing.values())
    # print(lab_file_listing.items())
    #
    # print(*lab_file_listing.keys())
    # print(*lab_file_listing.values())
    # print(*lab_file_listing.items())

    # print(list(lab_file_listing.keys()))
    # print(list(lab_file_listing.values()))
    # print(list(lab_file_listing.items()))

# -------------------------collection of abnormal laboratory data


print('----------------------<<< Laboratory data   >>> --------------------------------')

gender = 'm'

#----------------------------------------- Excel file extraction

import xlrd
file_location = 'C:/GDSRC/KOJ.xls'
workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_name('Sheet1')
KOJ_excel = worksheet._cell_values
# print("-"*75)
# -----------------------------------starting Excel selection items
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

# ---------------------------------- data change reference < M: >
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

# -------------------------------------<  common  > change items
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

# --------------------------------------<  female  > change items
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

# ----------------------------------------reference high vs low
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

# ------------------------------------------------------printing
    if row[0] != "AAA":
        print(row[0].ljust(7), row[1].ljust(25), row[2].rjust(12),
              "  ", row[3].ljust(15), row[4].rjust(10),"   ",
              row[5].ljust(10))
        # --------------------------------------------file_saving_definition
        cbc_lft = (row[0].ljust(7), row[1].ljust(25), row[2].rjust(12),
                   "  ", row[3].ljust(15), row[4].rjust(10), "   ",
                   row[5].ljust(10))  # ----------------save result

        with open('C:/GDSRC/RC_result_file.txt', 'a', encoding='utf-8') as f:
                f.writelines('\n')
                f.writelines(cbc_lft)

        # --------------------------------------------file_saving_definition
print("-"*75)
with open('C:/GDSRC/RC_result_file.txt', 'a', encoding='utf-8') as f:
    f.writelines('\n')
    f.writelines("-"*75)

# -------------------------------------character laboratory data
for row in (KOJ_excel[1:64]):
    keyword = row[0]
    trow = (row)
    trow[1].rstrip()
    row = trow
    if row[0] == 'AAA':
        del row[0]
        print(row[0].ljust(5), row[1].ljust(21), row[2].rjust(18),
              "  ", row[3].ljust(10),row[4].ljust(18))
        # --------------------------------------------file_saving_definition
        abo_hbv = (row[0].ljust(5), row[1].ljust(21), row[2].rjust(18),
              "  ", row[3].ljust(10),row[4].ljust(18))  # ----------------save result
        with open('C:/GDSRC/RC_result_file.txt', 'a', encoding='utf-8') as f:
                f.writelines('\n')
                f.writelines(abo_hbv)
        # --------------------------------------------file_saving_definition
print("-"*75)
with open('C:/GDSRC/RC_result_file.txt', 'a', encoding='utf-8') as f:
    f.writelines('\n')
    f.writelines("-"*75)
#
#          # ------------------------------------------urine laboratory data
for row in (KOJ_excel[64:87]):
    keyword = row[0]
    trow = (row)
    del (trow[0:2])
    del (trow[1])
    del (trow[3:5])
    del (trow[5:])
    trow[1].rstrip()
    row = trow
    print(row[0].ljust(7), row[1].ljust(25), row[2].rjust(12),
          "  ", row[3].ljust(15), row[4].rjust(10), "   ",)
#     # --------------------------------------------file_saving_definition
    urine = (row[0].ljust(7), row[1].ljust(25), row[2].rjust(12),
          "  ", row[3].ljust(15), row[4].rjust(10), "   ",)
    with open('C:/GDSRC/RC_result_file.txt', 'a', encoding='utf-8') as f:
        f.writelines('\n')
        f.writelines(urine)
#     # --------------------------------------------file_saving_definition
#
# print("-"*75)
# with open('C:/GDSRC/RC_result_file.txt', 'a', encoding='utf-8') as f:
#     f.writelines('\n')
#     f.writelines("-"*75)

