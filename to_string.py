import pickle
import pandas as pd
import numpy as np
import xlrd
import os
from K_main_mymodules import Fa_pre,Fa_rwb,Fa_val

FS = Fa_rwb.file_save
Fa_pre.status_check()
# -------------------------------------------------------------------------file saving definition
def ref_split(row):
    return row.split(':')[-1]

def ref_split_01(row):
    return row.split(i)[-1]

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
lab_res = pd.read_excel(Fa_val.br_ex_20 + 'D_2020_' +
                        name + '_.xlsx', header=0, index=False)
lab_C_selection = lab_res[['수진자명', '검사명', '검사결과', '단위', '참고치', 'dt_code_1']]
Lcs = lab_C_selection
print(Lcs)

df1 = Lcs[Lcs['검사명'].str.contains('ABO')]
ABO1 = df1.iloc[0, 2]
df2 = Lcs[Lcs['검사명'].str.contains('Rh')]
Rh1 = df2.iloc[0, 2]

# if gender == "f":
#     for i, ii in [('Uric','2.8-6.1'),('GGT','0-38'),('CK','34-145'),('Iron','32-153'),
#                 ('TIBC','223-422'),('RBC','3.70-5.2'),('Hemoglobin','11.0-16.0'),
#                 ('Hematocrit','36.0-46.0'),('WBC','4.00-11.0'),('ESR','0-20')]:
#         df2.loc[df2['검사명'].str.contains(i), 'ref_2'] = ii


#     if row[0] != "AAA":
#         srow = []
#         srow = (row[7].split("-"))
#         srow[0] = float(srow[0])
#         srow[1] = float(srow[1])

#         value1 = (float(srow[1])) - (float(row[3]))
#         value2 = (float(srow[0])) - (float(row[3]))

#         if (100 + value1 < 100):   row.append("High")
#         elif (100 + value2 > 100): row.append("Low")
#         else:                      row.append(".")


print (df2)


