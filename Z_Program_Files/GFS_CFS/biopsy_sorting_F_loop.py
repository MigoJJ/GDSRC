import pandas as pd
import numpy as np

path_b = ('C:/GDSRC/Data/source/')
b_items = pd.read_excel(path_b +'April_20_Lab.xlsx')
result_biopsy = b_items[['수진자명', '검사명','검체명','보험코드','검사결과']]
M = result_biopsy.sort_values(['보험코드'])

b_name = '이형종'
MigoJJ = pd.DataFrame()

for i in ['C5602', 'C5603008', 'C5621008']:
    print(i)
    MigoJJ = M.loc[M['보험코드'].isin([i]) & (M['수진자명'].isin([b_name]))]
    isempty = MigoJJ.empty
    if isempty == True:
        print('Is the DataFrame empty :', isempty)
        continue
    else:
        print(MigoJJ)
        print(type(MigoJJ))

        MigoJJ.to_csv(path_b + 'Biopsy_result.txt', mode = 'a')