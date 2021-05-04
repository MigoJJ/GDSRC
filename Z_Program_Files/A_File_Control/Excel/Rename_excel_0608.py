import os
import glob
from shutil import copyfile
import shutil

print(os.getcwd())
os.chdir('C:/Users/gdclinic001/Desktop/2019/labgen_2019/E_2019/')
for data_list in os.listdir():
    print(data_list)
    path_t = 'C:/Users/gdclinic001/Desktop/2019/labgen_2019/E_2019/'
    # os.chdir(path_t + './E_2019_01/')
    print(os.getcwd())
    data_list_1 = data_list[-8:]

    copyfile(data_list, './GDS/D_2019_' + data_list_1)
