import pandas as pd
import os
import subprocess

subprocess.call(['cls'], shell=True)
os.chdir('C:/GDSRC')
print(os.getcwd())
df = pd.read_excel('./Data/source/Main20.xlsx')

df = df.sort_values(by = ['성명'])
df = df.sort_values(by=['예진일_DATE9'], ascending=False)

df = df[~df['성명'].isnull()]
df = df[~df['검진종류'].isnull()]

# print(df)

df_2020 = df[df['성명'].str.endswith('20')]
df_2020 = df_2020.iloc[2:17,:]
df_2020.to_excel('./Data/source/Main20_07.xlsx', index=False)
print(df)


# for i in ['19887','19833','19672','19921','20102','20186']
#     df = df[df['검사번호'] != i]
# print(df)

# duplicate = df[df['예진일_DATE9'].duplicated(keep='last')]
# print(duplicate[['성명','예진일_DATE9']])

# df['성명'] = df['성명'].str[:-7]
# print(df['성명'])
# print(df)

