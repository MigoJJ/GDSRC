# Main GDSRC Directory Making

# import os
# os.rmdir('C:/GDSRC')
# try:
    # os.mkdir('C:/GDSRC')
    # os.mkdir('C:/GDSRC/Ref_Text')
    # os.mkdir('C:/GDSRC/Data')
    # os.mkdir('C:/GDSRC/Result')
    #
    # print(os.getcwd())
    # print(os.path.isdir('C:/GDSRC'))
    # path = os.path.join('C:/GDSRC','Text_File.txt')
#
# except Exception as i:
#     print(i)
import os
import shutil
try:
    def dir_making(x):
        for k in ('C:/GDSRC/Data/','C:/GDSRC/Result/'):  # 'C:/GDSRC/Ref_Text'
            root_path = k
            folders = []
            # print(folders)
            folders.append(x)
            for folder in folders:
                os.mkdir(os.path.join(root_path,folder))

    dir_making(input('화일 이름을 넣으세요 ... :이름_YY_MM_DD_  = '))
except Exception as i:
    print(i)

# GDSRC Back-up



#  BU = " "
#  input("... Back-up 하시겠습니까 ? " + BU)
#  if BU == 'y':
# import shutil
# import datetime as d
# import os
#
# koh =d.datetime.now()
# ddate = (str(d.datetime.now()))
# Name = 'GDS_BU'
# try:
#     i = ddate.split(' ')
#     jae = i[0].split('-')
#     joon = '_'.join(jae)
#     print(jae)
#     print(joon)
#
#     MD = (joon,'_{}'.format(Name))
#     # shutil.rmtree('C:/GDS_Python/Copy_File_Control')
#     shutil.copytree('C:/GDSRC', 'C:/BU/GDSRC_BU_{}'.format(joon))
#     # os.mkdir('c:/{}'.format(MD))
# except Exception as i:
#     print(i)