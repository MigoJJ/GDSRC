# ......... .... .... .... .... .... .... .... .... .... .... ......
# .... GDS Clinic Routine Check Management Program version 1.30 ....
# ......... .... .... .... .... .... .... .... .... .... .... ......
import sys
import pickle
import subprocess
sys.path.append('C:/GDSRC/K_main_mymodules')
from K_main_mymodules import Fa_pre,Fa_rwb
real_data_dict = {}
today_nl = []
nl_main = []
nl_lab = []
Fa_pre.status_check()
# --------------------------------------------------------------name listing
df = Fa_rwb.Loader_M()
df_L = Fa_rwb.Loader_L()
df_L.drop_duplicates(subset='수진자명', keep='first', inplace=True)
nl_lab = df_L['수진자명'].tolist()

df_M = Fa_rwb.Loader_M()
df_M.drop_duplicates(subset='성명', keep='first', inplace=True)
nl_main = df_M['성명'].tolist()

for l in nl_main:
    print(' l   :   ',l)
    for m in nl_lab:
        if m in l:
            today_nl.append(l)
print('Lab data number   :   ',len(nl_lab),nl_lab)
print('Main data number   :   ',len(nl_main),nl_main)
print('Sorted data number :   ',len(today_nl),today_nl)
name_list = today_nl
# ---------------------------------------------------pickling and batch file
# name_list = ['홍금아0715-20']
# ---------------------------------------------------pickling and batch file
for Rhimm in name_list:
    df1 = df[df['성명'] == Rhimm]
    check_date = df1.iloc[0, 10]
    C_date = check_date.strftime('%Y_%m_%d')
    gender = df1.iloc[0, 7]
    real_data_dict = {'R_name': Rhimm, 'R_date': C_date, 'R_gender': gender}
    dictionary_data = real_data_dict

    print(real_data_dict)

    with open("C:/GDSRC/data.pkl", "wb") as a_file:
        pickle.dump(dictionary_data, a_file)
    a_file.close()

    subprocess.call(['cls'], shell=True)
    subprocess.call(['C:/GDSRC/1.bat'], shell=True)
    subprocess.call(['cls'], shell=True)
    real_data_dict = {}