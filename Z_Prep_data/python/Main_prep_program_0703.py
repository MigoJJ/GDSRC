import pandas as pd
import os
import datetime as dt
from datetime import datetime

os.chdir('C:/GDSRC')
print(os.getcwd())
df = pd.read_excel('./Data/source/Main20_07_20.xlsx')
df = df[['고객 ID','성명','예진일_DATE9']]
df = df.fillna('blank')

name_split = df['성명'].str.split('-')
df["ft_name"] = name_split.str.get(0)
df["year"] = name_split.str.get(1)

df = df[df['성명'] != 'blank']

df.loc[df['성명'] == '김성호0608','성명'] = '김성호0608-19'

df = df[df['성명'].str.contains('20')]