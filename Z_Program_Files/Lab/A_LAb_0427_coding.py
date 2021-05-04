#--------------------------------------------------------------import pickle data
import xlrd
import pickle
a_file = open("data.pkl", "rb")
output = pickle.load(a_file)
print(output)
a_file.close()

name = (output['D_name'])
gender = (output['D_gender'])
age = (output['D_age'])
print(name,gender,age)
#--------------------------------------------------------------file saving definition
def file_save(x):
    with open('C:/GDSRC/Result/RC_result_file.txt', 'a',encoding='utf-8') as f:
        for item in (x):
            f.writelines(item)
            # f.writelines("%s\n" % item)
    f.close()

file_save("*********************************************************\n\n")
file_save("              GDS Clinic Laboratory Result       \n\n")
file_save("*********************************************************\n\n")
#--------------------------------------------------------------Excel file extraction

name = name[:3]
file_location = ('C:/GDSRC/Data/' + 'Result_lab_' + name + '.xlsx')  #-------data excel file name


workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_name('Sheet1')
KOJ_excel = worksheet._cell_values

for row in (KOJ_excel[1:64]):
    keyword = row[0]
    # -------------------------------------- << INSERT lAB DATA REFERENECE CHANGE >> ---

    lab_file_listing = {}
    lab_file_listing = dict({'code': row[5], 'value': row[6], 'unit': row[9], 'reference': row[10]})
    # print(lab_file_listing)
    list_lab = list(lab_file_listing.values())

    # -----------------------------------------------------ABO  Rh
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
        # -------------------------------------------------Hepatitis
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
                      '''국내에서 사용 중인 A형 간염 백신은 병원체를 배양한 후 포르말린으로 불활성화시킨 백신으로 사백신이라고도 불
리는 불활성화 백신에 속합니다. 살아있지 않기 때문에 몸 안에서 증식할 수 없으며, 면역력이 약한 사람에게 
투여해도 감염증을 유발할 수 없어 안전하다. 대신 1회 접종으로는 완벽한 예방이 어렵고 장기적인 면역력
이 생성되기 위해서는 추가 접종이 필요합니다. 2회 접종 시 94~100%의 예방 효과를 보이며, 장기간의 효과
에 대한 연구는 아직 없으나 이론적으로 성인에서는 25년 이상, 소아에서는 14~20년간 예방 효과가 유지될 
것이라 예측되고 있습니다.
 \n''')

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
            file_save('C 형 간염이 없습니다.\n\n')
        else:
            print('C 간염의 항체(Ab ,Antibody)가 있습니다. \n  C 형 간염의 확정적인 검사는 아닙니다. 전문의의 진료를 받도록하십오.\n')
            file_save('C 간염의 항체(Ab ,Antibody)가 있습니다. \n  C 형 간염의 확정적인 검사는 아닙니다. 전문의의 진료를 받도록하십오.\n')


file_save('<<   Blood Chemistry data   >> \n')
file_save('-'*75 + '\n')

# ----------------------------------------- Excel file extraction
import xlrd
file_location = ('C:/GDSRC/Data/' + 'Result_lab_' + name + '.xlsx')  #-------data excel file name
workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_name('Sheet1')
KOJ_excel = worksheet._cell_values

abn_val = []

for row in (KOJ_excel[1:64]):
    keyword = row[0]
    trow = (row)
    del (trow[0:2])
    del (trow[0])
    del (trow[1])
    del (trow[3])
    del (trow[6:])
    trow[1].strip()
    row = trow
    # print(*row)
# ----------------------------------- data change reference <character >
    def charact_lab(x):
        if row[1].startswith(x):
            row.insert(0,"AAA")
            crow = row.copy()
            return crow
    for i in ('ABO', 'Rh', 'RF', 'RPR', 'HIV', 'HAV', 'HCV', 'HBs','Diff'):
        charact_lab(i)

# ----------------------------------- data change reference < M: 성인: F:>
    def change_ref (x):
        if row[5].startswith(x):
            if row[5].startswith('M:'):
                row.insert(5, row[5][2:])  # List 4 item  from 2 character
            if row[5].startswith('F:'):
                row.insert(5, row[5][2:])  # List 4 item  from 2 character
            if row[5].startswith('성인'):
                row.insert(5, row[5][3:])  # List 5 item  from 2 character
            row.pop()
            return row
    change_ref('M:')
    change_ref('F:')
    change_ref('성인')

# ------------------------------common reference replace   < \n >
# -------------------------------------<  common  > change items
    def change_items (x,y):
        if row[1].startswith(x):
            row.insert(5,y)
            del row[6]
            return row
    for i,ii in [('BUN','8.0-23.0'),('Amylase','28-110'),('Cholesterol','110-200'),('Tri','0.0-150'),
                 ('HDL','0-40'),('LDL','0-100'),('Glucose','70-100'),('25-OH','20-100'),('AFP','0.0-8.1'),
                 ('CEA','0.0-5.0'),('CA19-9','0.0-37.0'),('CA125','0.0-35.0'),('PSA','0.0-4'),('ESR','0.0-15'),
                 ('GGT', '0.0-73.0'),('ALP', '35-130'),('Sodium','135-145'),('Chloride','98-110'),('WBC','4.00-11.00'),
                 ('MCV','80-100'),('MCH','26.0-34.0'),('Other','0-1'),('RBC', '4.5-5.9'),
                 ('Hemoglobin', '12.5-17.5'),('Hematocrit', '38.0-53.0')]:
        change_items(i,ii)

    if row[2].startswith('<'):
        print('Migo Good~~!')
        row.insert(2, row[2][2:])  # List 4 item  from 1 character
        print(row[2])
# --------------------------------------<  female  > change items
    if gender == "f":
        for i, ii in [('Uric','2.8-6.1'),('GGT','0-38'),('CK','34-145'),('Iron','32-153'),
                      ('TIBC','223-422'),('RBC','3.70-5.2'),('Hemoglobin','11.0-16.0'),('Hematocrit','36.0-46.0'),
                      ('WBC','4.00-11.0'),('ESR','0-20')]:
            change_items(i, ii)
 ## ----------------------------------------reference high vs low
    if row[0] != "AAA":
        srow = []
        # print(*row[0:6])
        srow = (row[5].split("-"))
        # row[2] = float(row[2])

        value1 = (float(srow[1])) - (float(row[2]))
        value2 = (float(srow[0])) - (float(row[2]))
        if (100 + value1 < 100):
            row.append("High")
        elif (100 + value2 > 100):
            row.append("Low")
        else:
            row.append(".")

    row[2] = str(row[2])
# -------------------------------------------------------printing
    if row[0] != "AAA":
        print (row[0].ljust(7),row[1].ljust(25),row[2].rjust(10),row[3].ljust(15), row[4].rjust(7),row[5].rjust(10), row[6].rjust(7))
        # --------------------------------------------file_saving_definition
        cbc_lft = (row[0].ljust(7), row[1].ljust(25), row[2].rjust(10),
                   row[3].ljust(15), row[4].rjust(10),
                   row[5].rjust(10)+row[6].rjust(7)+'\n')  # ----------------save result
        file_save(cbc_lft)

        if row[6] != ".": # ---------------------------------abnormal values
            abn_val.append(list(row))
        # --------------------------------------------file_saving_definition
print("-"*75)
with open('C:/GDSRC/Result/RC_result_file.txt', 'a', encoding='utf-8') as f:
    f.writelines('\n')
    f.writelines("-"*75)
    f.writelines('\n')
# -------------------------------------character laboratory data
for row in (KOJ_excel[1:64]):
    keyword = row[0]
    trow = (row)
    trow[1].rstrip()
    row = trow
    if row[0] == 'AAA':
        del row[0]
        if row[1] == 'Differential count':
            continue
        print(row[0].ljust(5), row[1].ljust(21), row[2].rjust(18),
              "  ", row[3].ljust(10),row[4].ljust(18))
        # --------------------------------------------file_saving_definition
        abo_hbv = (row[0].ljust(5), row[1].ljust(21), row[2].rjust(18),
              "  ", row[3].ljust(10),row[4].ljust(18)+'\n')  # ----------------save result
        file_save(abo_hbv)
        # --------------------------------------------file_saving_definition
print("-"*75+'\n')
with open('C:/GDSRC/Result/RC_result_file.txt', 'a', encoding='utf-8') as f:
    f.writelines('\n')
    f.writelines("-"*75)
    f.writelines('\n')
# ------------------------------------------urine laboratory data
print("-" * 75 + '\n')
file_save("-" * 75 + '\n')
file_save("  << Urine Chemical Test  >>  \n")

for row in (KOJ_excel[65:91]):
    keyword = row[0]
    trow = (row)
    del (trow[0:2])
    del (trow[0])
    del (trow[1])
    del (trow[3])
    del (trow[6:])
    trow[1].strip()
    row = trow

    if row[1] == '요침사검사':
        print("-" * 75 + '\n')
        file_save("-" * 75 + '\n')
        file_save("  << Urine Microscopic Test  >>  \n")
        continue
    if row[1] ==  'Microscopy(S)  ':
        print("-" * 75 + '\n')
        file_save("-" * 75 + '\n')
        file_save("  << Stool Test  >>  \n")
        continue
    if row[1] ==  '일반 부인과 검사(Pap stain, GY)':
        print("-" * 75 + '\n')
        file_save("-" * 75 + '\n')
        file_save("  << PAP Stains  >>  \n\n")


    print(row[0].ljust(7), row[1].ljust(25), row[2].rjust(12),
          "  ", row[3].ljust(15), row[4].rjust(10), "   ",)
    # --------------------------------------------file_saving_definition
    urine = (row[0].ljust(7), row[1].ljust(25), row[2].rjust(12),
          "  ", row[3].ljust(15), row[4].rjust(10), "\n",)
    file_save(urine)
    # --------------------------------------------file_saving_definition
print("-"*75 +'\n')
with open('C:/GDSRC/RC_result_file.txt', 'a', encoding='utf-8') as f:
    f.writelines('\n')
    f.writelines("-"*75)
# -----------------------------------------collection of abnormal laboratory data
with open('C:/GDSRC/Result/RC_result_file.txt', 'a', encoding='utf-8') as f:
    f.writelines('\n')
    f.writelines("  <<  Abnormal Lacoratory data  >>  ")
    f.writelines("-"*37)
    f.writelines('\n')


for w in range(0, (len(abn_val))):
    # print(abn_val)
    abn_v = (abn_val[w])
    abn_va = (abn_v[1].ljust(25), abn_v[2].rjust(10),
          abn_v[3].ljust(10), abn_v[4].rjust(8),
          abn_v[5].rjust(10))
    print(abn_va)
    file_save(abn_va)
    file_save('\n')

    with open('C:/GDSRC/RC_result_file.txt', 'a', encoding='utf-8') as f:
        f.writelines('\n')
        f.writelines(abn_va)
