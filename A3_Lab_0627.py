# ......... .... .... .... .... .... .... .... .... .... .... ......
# .... GDS Clinic Routine Check Management Program version 1.30 ....
# ......... .... .... .... .... .... .... .... .... .... .... ......
import pickle
import pandas as pd
import xlrd
from K_main_mymodules import Fa_pre,Fa_rwb,Fa_val
FS = Fa_rwb.file_save
Fa_pre.status_check()
# -------------------------------------------------------------------------file saving definition
def line_sub(x, y):
    if row[2].startswith(x):
        FS('-' * 75 + '\n')
        FS(y)

def hepatitis_report(L, C):
    hep_res = hep.iloc[L, C]
    FS('\n Hepatitis : ' + hep_res + '\n')
# -----------------------------------------------------------------------------import pickle data
abn_val = []
hep_list = []

a_file = open('C:/GDSRC/data.pkl', "rb")
output = pickle.load(a_file)
a_file.close()

name = (output['R_name'])
name = name[:-7]
gender = (output['R_gender'])
age = (output['R_age'])
# --------------------------------------------------------------------------Excel file extraction
Fa_rwb.sa_subtitle()

lab_res = pd.read_excel(Fa_val.br_ex_20 + 'D_2020_' + name + '_.xlsx', header=0, index=False)
lab_C_selection = lab_res[['수진자명', '검사명', '검사결과', '단위', '참고치','dt_code_1']]
Lcs = lab_C_selection

file_location = (Fa_val.br_ex_20 + 'D_2020_' + name + '_.xlsx')  # -----------------------file name
workbook = xlrd.open_workbook(file_location)
ws_lab = workbook.sheet_by_name('Laboratory')
ws_tis = workbook.sheet_by_name('Tissue')
ws_PAP = workbook.sheet_by_name('PAP')
ws_uri = workbook.sheet_by_name('Urine')
ws_sto = workbook.sheet_by_name('Stool')

J_lab = ws_lab._cell_values
J_tis = ws_tis._cell_values
J_PAP = ws_PAP._cell_values
J_uri = ws_uri._cell_values
J_sto = ws_sto._cell_values

l_sheet_count = ws_lab.nrows
# print(Lcs.tail(5))
# -----------------------------------------------------------------------------------------ABO  Rh
df1 = Lcs[Lcs['검사명'].str.contains('ABO')]
ABO1 = df1.iloc[0, 2]
df2 = Lcs[Lcs['검사명'].str.contains('Rh')]
Rh1 = df2.iloc[0, 2]

FS('당신의 혈액 형은  {1} 형 ,  Rh {0} 입니다.\n'
          '응급실에서 수혈이 필요한 경우 필요한 혈액형을 신속하게 판단할 수 있는 중요한 정보 입니다.\n'.format(Rh1, ABO1))
# ----------------------------------------------------------------------------------------Hepatitis
df = pd.DataFrame(Lcs, columns=['결과완료일','수진자명', '검사명', '검사결과', '단위', '참고치'])

df.loc[df['검사명'].str.startswith('HAV') & df['검사결과'].str.contains('Pos'), 'hepatitis'] = ('1')
df.loc[df['검사명'].str.startswith('HAV') & df['검사결과'].str.contains('Neg'), 'hepatitis'] = ('2')
df.loc[df['검사명'].str.startswith('HBsAg') & df['검사결과'].str.contains('Pos'), 'hepatitis'] = ('3')
df.loc[df['검사명'].str.startswith('HBsAg') & df['검사결과'].str.contains('Neg'), 'hepatitis'] = ('4')
df.loc[df['검사명'].str.startswith('HBsAb') & df['검사결과'].str.contains('Pos'), 'hepatitis'] = ('5')
df.loc[df['검사명'].str.startswith('HBsAb') & df['검사결과'].str.contains('Neg'), 'hepatitis'] = ('6')
df.loc[df['검사명'].str.startswith('HCV') & df['검사결과'].str.contains('Pos'), 'hepatitis'] = ('7')
df.loc[df['검사명'].str.startswith('HCV') & df['검사결과'].str.contains('Neg'), 'hepatitis'] = ('8')

df = df.fillna(value='blank')
hep_list = df['hepatitis'].tolist()
remove_list = ['blank']
hep_list = [i for i in hep_list if i not in remove_list]

hep = pd.read_excel(Fa_val.br_la_rf, sheet_name='Charact', index_col=None, header=None)
# -------------------------------------------------------------------------------------------- HAV
if '1' in hep_list:   hepatitis_report(5, 1)
if '2' in hep_list:   hepatitis_report(5, 2)
# -------------------------------------------------------------------------------------------- HBV
if '3' in hep_list and '5' in hep_list:  hepatitis_report(6, 1)
if '3' in hep_list and '6' in hep_list:  hepatitis_report(7, 1)
if '4' in hep_list and '5' in hep_list:  hepatitis_report(8, 1)
if '4' in hep_list and '6' in hep_list:
    if age > 60:
        hepatitis_report(9, 1)
    else:
        hepatitis_report(9, 2)
# ------------------------------------------------------------------------------------------- HCV
if '7' in hep_list:        hepatitis_report(10, 1)
if '8' in hep_list:        hepatitis_report(10, 2)
# -------------------------------------------------------------------------------  Laboratory Data
FS('-' * 75 + '\n')
FS('<<   Blood Chemistry data   >> \n')

for row in (J_lab[1:l_sheet_count]):
    keyword = row[0]
    if row[2].startswith('CBC'):
        continue

    line_sub('Uric acid', '        ___Muscle Enzyme___ \n')
    line_sub('S_Glucose', '        ___Pancreas Enzyme___ \n')
    line_sub('BUN', '        ___Reanl Function___ \n')
    line_sub('Cholesterol', '        ___lipid battery Function___ \n')
    line_sub('WBC count', '<<   Complete Blood Counts Data   >> \n')
    line_sub('Lymphocyte', '        ___WBC Differential count___ \n')
    line_sub('AFP', '<<   Tumor Markers   >> \n')
    line_sub('25-OH-Vit.D', '<<   Vitamin D   >> \n')
    line_sub('TSH', '<<   Thyroid Function Test   >> \n')

    # -------------------------------------------------------- data change reference <character >
    def change_items(x, y):
        if row[2].startswith(x):
            row.insert(7, y)
            del row[8]
            return row

    def charact_lab(x):
        if row[2].startswith(x):
            row.insert(0, "AAA")
            crow = row.copy()
            return crow

    for i in ('ABO', 'Rh', 'RF', 'RPR', 'HIV', 'HAV', 'HCV', 'HBs', 'HBV',
              'Diff','LH','FSH','E2','Measles','HBe','Thyroglobulin','Anti-TPO',
              'TSH-R-Ab','Prolactin','IGF','Growth','Cortisol'):
        charact_lab(i)
        # ---------------------------------------------------data change reference < M: F: 성인:>
    def change_ref(x, y):
        row[7] = str(row[7])
        if row[7].startswith(x):
            row.insert(7, row[7][y:])  # -------------------------- List 4 item  from 2 character
            row.pop()
            return row

    change_ref('M:', 2)
    change_ref('F:', 2)
    change_ref('성인', 3)

    for i,ii in [('BUN','8.0-23.0'),('Amylase','28-110'),('Cholesterol','110-200'),('Tri','0.0-150'),
                 ('HDL','0-40'),('LDL','0-100'),('S_Glucose','70-100'),('25-OH','20-100'),('AFP','0.0-8.1'),
                 ('CEA','0.0-5.0'),('CA19-9','0.0-37.0'),('CA125','0.0-35.0'),('PSA','0.0-4'),('ESR','0.0-15'),
                 ('GGT', '0.0-73.0'),('ALP', '35-130'),('Sodium','135-145'),('Chloride','98-110'),
                 ('WBC','4.00-11.00'),('MCV','80-100'),('MCH','26.0-34.0'),('Other','0-1'),('RBC', '4.5-5.9'),
                 ('Hemoglobin', '12.5-17.5'),('Hematocrit', '38.0-53.0'),('Cyfra','0-3.30'),
                 ('Cortisol','3.44-22.45'),('Free testosterone','5.40-40.00')]:
        change_items(i,ii)

    if row[3].startswith('<'):
        row.insert(3, row[3][2:])
        del(row[4])
    #
    if ':' in row[3]:               # ---------------------------------------result data trimming
        s_row = row[3].split(':')
        s_row.pop()
        row[3] = float(*s_row)
    #         # --------------------------------------------------------<  female  > change items
    if gender == "f":
        for i, ii in [('Uric','2.8-6.1'),('GGT','0-38'),('CK','34-145'),('Iron','32-153'),
                      ('TIBC','223-422'),('RBC','3.70-5.2'),('Hemoglobin','11.0-16.0'),
                      ('Hematocrit','36.0-46.0'),('WBC','4.00-11.0'),('ESR','0-20')]:
            change_items(i, ii)
                                # ----------------------------------------- reference high vs low
    if row[0] != "AAA":
        srow = []
        srow = (row[7].split("-"))
        srow[0] = float(srow[0])
        srow[1] = float(srow[1])

        value1 = (float(srow[1])) - (float(row[3]))
        value2 = (float(srow[0])) - (float(row[3]))

        if (100 + value1 < 100):   row.append("High")
        elif (100 + value2 > 100): row.append("Low")
        else:                      row.append(".")
    row[2] = str(row[2])  # ----------------------------------------------------------- STR FLOAT
    if row[0] != "AAA":
        d_dat = str(row[0])
        cbc_lft = (d_dat[3:8], row[2].rjust(19) + str(row[3]).rjust(12),'     ',
                   row[7].ljust(18), row[6].ljust(12) + row[9] +  '\n')
        FS(cbc_lft)
        if row[9] != ".":  # -----------------------------------------------------abnormal values
            abn_val.append(list(row))

Fa_rwb.liner()
# ----------------------------------------------------------------------character laboratory data
for row in (J_lab[1:l_sheet_count]):
    keyword = row[0]
    trow = (row)
    row = trow

    if row[0] == 'AAA':
        del row[0]
        if row[3].startswith('Lymphocyte'):
            continue
        line_sub('AB', '<<   Blood Typing   >> \n')
        line_sub('HAV-Ab', '<<  Hepatitis  >>\n')
        line_sub('RF 정성', '<<  Inflammation Test  >>\n')

        d_dat = str(row[0])
        abo_hbv = (d_dat[3:8].ljust(15), row[2].ljust(25), row[3].ljust(12) + '\n')
        FS(abo_hbv)
# --------------------------------------------------------------------------urine laboratory data
FS("-" * 75 + '\n')
FS("  << Urine Chemical Test  >>  \n")

l_sheet_count = ws_uri.nrows

for row in (J_uri[1:l_sheet_count]):
    keyword = row[0]
    if row[2].startswith('Urine 10종'):
        continue
    elif row[2].startswith('요침사검사'):
        FS("  ___ Urine Microscopic Test ___   \n")
        continue
    else:
        d_dat = str(row[0])
        urine = (d_dat[3:8].ljust(15), row[2].ljust(25),
                 "  ", row[3].ljust(15), row[4].rjust(10), "\n",)
        FS(urine)

l_sheet_count = ws_sto.nrows
# -------------------------------------------------------------------------------------------stool
FS("-" * 75 + '\n')
FS("  << Stool Test  >>  \n")
for row in (J_sto[1:l_sheet_count]):
    keyword = row[0]
    if row[2].startswith('Microscopy'):
        continue
    else:
        d_dat = str(row[0])
        stool = (d_dat[3:8].ljust(15), row[2].ljust(25),"  ", row[3].ljust(15), "\n",)
        FS(stool)
# --------------------------------------------------------collection of abnormal laboratory data
Fa_rwb.liner_ab()

for w in range(0, (len(abn_val))):
    abn_v = (abn_val[w])
    abn_va = (abn_v[1].ljust(15),
              abn_v[2].ljust(20), str(abn_v[3]).ljust(10),abn_v[7].rjust(15),abn_v[9].rjust(10))
    FS(abn_va)
    FS('\n')
