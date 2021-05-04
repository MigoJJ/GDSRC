import os
import pandas as pd

os.chdir('C:/MigoJJ')
print(os.getcwd())

df = pd.read_excel('./지디스내과(NEW)(20190101-20190630).xlsx')
df = df.fillna('blank')
# df = df[df['보험코드'].str.contains('C5602')]
df = df[df['수진자명'] == '엄익부']
# df = df.drop_duplicates(['수진자명'], keep='first')
print(df.head(30))
print(df.iloc[:,[2,4,5,6]][60:90])


