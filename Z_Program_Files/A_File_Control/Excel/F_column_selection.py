# Importing common libraries
import pandas as pd
from pandas import DataFrame
import numpy as np

df = pd.read_excel('C:/GDSRC/Ref_Text/April_20_Lab.xlsx')
# print(df)

print(df.columns.ravel())
# print(df['수진자명'].tolist())

print(df.iloc[1:10,[2,9]])

for i in df.iloc[1:10,[2,5]]:
    print(i +' dlsthogkqslek')
    














# df= pd.concat([df[df.columns[2]], df[df.columns[4:6]]], axis=1)
# print(df)
# 
# df= pd.concat([df[df.columns[0]]], df[df.columns[4:6]]], axis=1)
# print(df)

# print(named_selection['수진자명'])