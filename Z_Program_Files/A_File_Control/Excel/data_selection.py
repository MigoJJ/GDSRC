import pandas as pd
import os

os.getcwd()
os.chdir('C:/GDSRC/Data/source')
print(os.listdir())
# print(os.walk('C:/GDSRC/Data/source', topdown=False))

real_data_dict = {}
path = os.getcwd()
#-------------------------------------------name listing
# name_list = df['성명'].values.tolist()
