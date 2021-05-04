import pandas as pd
import numpy as np

path_b = ('C:/GDSRC/Data/source/')
b_items = pd.read_excel(path_b +'April_20_Lab.xlsx')
result_biopsy = b_items[['수진자명', '검사명','검체명','보험코드','검사결과']]
M = result_biopsy.sort_values(['보험코드'])
print(M)

MigoJJ = M.loc[M['보험코드'].isin(['C5602']) & (M['수진자명'].isin(['임정식']))]
print(MigoJJ.iloc[:,3:5])
MigoJJ.to_csv(path_b + 'sample.txt','a')

# MigoJJ = M.loc[M['보험코드'].isin(['C5603008']) & (M['수진자명'].isin(['권오상']))]
# print(MigoJJ.iloc[:,0:5])
# MigoJJ.to_csv(path_b + 'sample.txt','a')

# MigoJJ = M.loc[M['보험코드'].isin(['C5621008']) & (M['수진자명'].isin(['오갑선']))]
# print(MigoJJ.iloc[:,0:5])
# MigoJJ.to_csv(path_b + 'sample.txt','a')

