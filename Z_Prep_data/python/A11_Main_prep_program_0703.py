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
# print(df.head(30))
#
df = df[df['성명'] != 'blank']
# print(df.head(10))

df.loc[df['성명'] == '김성호0608','성명'] = '김성호0608-19'
# print(df.head(10))


df = df[df['성명'].str.contains('20')]
print(df.head(30))
print(len(df))
print(df)

# df.to_excel('C:/Exercise/2019_Main.xlsx')

# df["수진자명"]=df["성명"].str[:-7]
# df["date"]=df["성명"].str[-7:-3]
# df["year"]=df["성명"].str[-2:]
# df = df.fillna('blank')
# print(df.tail(6))

# df.fillna('blank')
# df1 = df[df['수진자명'] == 'blank']
# print(df1.head())

# https://www.it-swarm.dev/ko/python/%EC%A1%B0%EA%B1%B4%EB%B6%80-%ED%91%9C%ED%98%84%EC%8B%9D%EC%9D%84-%EA%B8%B0%EB%B0%98%EC%9C%BC%EB%A1%9C-pandas-dataframe%EC%97%90%EC%84%9C-%ED%96%89%EC%9D%84-%EC%82%AD%EC%A0%9C%ED%95%98%EB%8A%94-%EB%B0%A9%EB%B2%95/1070075551/
