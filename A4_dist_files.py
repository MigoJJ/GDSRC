# ......... .... .... .... .... .... .... .... .... .... .... ......
# .... GDS Clinic Routine Check Management Program version 1.30 ....
# ......... .... .... .... .... .... .... .... .... .... .... ......

import pickle
import os
import shutil
from K_main_mymodules import Fa_pre,Fa_val

Fa_pre.status_check()
# ---------------------------------------------------reading pickle
a_file = open('C:/GDSRC/data.pkl', "rb")
output = pickle.load(a_file)
a_file.close()

name   = (output['R_name'])
b_name = name[:-7]
gender = (output['R_gender'])
ID     = (output['R_ID'])
ID     = ID[:7]
s_date =(output['R_date'])
# # ---------------------------------------------------coping directory
B = (s_date + '__' + b_name + '__' + ID)
ex_file = (Fa_val.br_ex_20 + 'D_2020_' + b_name +'_.xlsx')

os.makedirs('./Result/RC_2020/' + B + '/',exist_ok=True)
BD = ('./Result/RC_2020/' + B + '/')
F_result  = (B + '_.txt')
FE_result = (B + '_.xlsx')

shutil.copy(Fa_val.br_re, BD + B + '_.txt')
shutil.copy(ex_file, BD + FE_result)
