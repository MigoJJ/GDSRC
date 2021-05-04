import os
import numpy
import pandas as pd

cwd = os.path.abspath('')
files = os.listdir(cwd)

files = os.listdir(cwd)
set_files = files[1:4]

print (files)
print (set_files)

df = pd.DataFrame()
for file in set_files:
     if file.endswith('.xlsx'):
         df = df.append(pd.read_excel(file), ignore_index=True)
df.head()

print (df)