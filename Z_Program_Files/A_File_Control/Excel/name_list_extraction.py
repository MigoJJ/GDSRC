import pandas as pd 

df = pd.read_excel('C:/GDSRC/Data/source/Main_20_04.xlsx')
# ---  from Access  GDSRC  Main  Sheet  --- 

df = df.iloc[:,6]
print(df)

df.to_csv('C:/GDSRC/name_list_2020.txt')