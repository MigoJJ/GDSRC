import shutil
import datetime as d
import os

dir = 'D:/GDSRC_Copy'
if os.path.exists(dir):
    shutil.rmtree(dir)

shutil.copytree('C:/GDSRC','D:/GDSRC_Copy')

koh =d.datetime.now()
ddate = (str(d.datetime.now()))
Name = 'GDS_BU'
try:
    i = ddate.split(' ')
    jae = i[0].split('-')
    joon = '_'.join(jae)
    print(jae)
    print(joon)

    MD = (joon,'_{}'.format(Name))
    # shutil.rmtree('C:/GDS_Python/Copy_File_Control')
    shutil.copytree('C:/GDSRC', 'C:/GDSRC_BU_{}'.format(joon))
    # os.mkdir('c:/{}'.format(MD))
except Exception as i:
    print(i)

# ------------------------------directory removal
