import os
import pandas as pd
import xlrd
from K_main_mymodules import Fa_pre,Fa_rwb,Fa_val
# _____________________________________________________ Reading data from excel
def Loader_M():
    os.chdir('C:/GDSRC/Data/source')
    df = pd.read_excel('./Main_0727.xlsx')
    df.fillna('blank')
    return df

def Loader_L():
    os.chdir('C:/GDSRC/Data/source/labgen_2020')
    df = pd.read_excel('./지디스내과(NEW)(20200713-20200727).xlsx')
    df.replace('^\s+', '', regex=True, inplace=True)  # front
    df.replace('\s+$', '', regex=True, inplace=True)  # end
    df.loc[df['보험코드'] == 'D3022003', '검사명'] = 'S_Glucose'
    df.fillna('blank')
    return df
# ___________________________________________ Labgenomics data parsing to excel
def file_save(x):
    with open('C:/GDSRC/Result/RC_result_file.txt', 'a', encoding='utf-8') as f:
        for item in (x):
            f.writelines(item)
    f.close()

def sa_title():
    file_save("*********************************************************\n\n")
    file_save("               GDS Clinic Routine Check Result           \n\n")
    file_save("*********************************************************\n\n")

def sa_subtitle():
    file_save("\n*********************************************************\n")
    file_save("              GDS Clinic Laboratory Result       \n")
    file_save("*********************************************************\n\n")

def liner():
    with open(Fa_val.br_re, 'a', encoding='utf-8') as f:
        f.writelines('\n')
        f.writelines("-" * 75)
        f.writelines('\n')

def liner_ab():
    with open(Fa_val.br_re, 'a', encoding='utf-8') as f:
        f.writelines('\n')
        f.writelines("\n\n\n\n\n___  <<  Abnormal Laboratory data  >> ___\n ")
        f.writelines('\n')

def pt(x, y):
    print(y)
    print(x + '\t\t\t' + '.'*40, type(y))
