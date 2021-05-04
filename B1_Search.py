import os
import shutil
import os.path
from os import path

Ter_File = os.listdir('./Result/RC_2020/')
Rer_File = os.listdir('./Result/Search_Patient/')

for i in Rer_File:
    # shutil.copytree('C:/GDSRC/Result/Search_Patient/' + i, 'C:/GDSRC/Result/Search_past_list_all/' + i)
    shutil.rmtree('./Result/Search_Patient/' + i)
# ---------------------------------------------------search files
s_name = input('\n\n찾을 사람의 이름을 넣어주세요   :   ...  ')
for i in Ter_File:
    print(i)
    if s_name in i:
        print('\n' + i)
        shutil.copytree('./Result/RC_2020/' + i, './Result/Search_Patient/' + i)

