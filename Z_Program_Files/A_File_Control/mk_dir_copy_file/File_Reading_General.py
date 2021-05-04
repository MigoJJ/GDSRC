# # 디렉토리 만들기
# import os
# print (os.getcwd())
# print(os.listdir(os.getcwd()))
#
# print(os.chdir('C:/GDS_Python'))
# print (os.getcwd())
# print(os.listdir(os.getcwd()))
#
# try:
#     # os.rmdir('C:/GDSRC')
#     os.mkdir('C:/GDSRC')
#     # path = os.path.join('C:/GDSRC','Text_File.txt')
#     print (os.getcwd())
#     print(os.path.isdir('C:/GDSRC'))
# except Exception as i:
#     print(i)

# 파일경로 새로 만들기
# creating src_dir, dst_dir
# base_dir = 'C:/Users/admin'
# src_dir = os.path.join(base_dir, 'src_dir')
# dst_dir = os.path.join(base_dir, 'dst_dir')
#
# os.mkdir(src_dir)
# os.mkdir(dst_dir)

# 디렉토리 카피와 제거
# import shutil
# shutil.rmtree('C:/GDS_Python/Copy_File_Control')
# shutil.copytree('C:/GDS_Python/File_Control','C:/GDS_Python/Copy_File_Control')
# shutil.rmtree('C:/GDS_Python/Copy_File_Control')

#파일 제거 os
# import os
# os.remove('C:/GDS_Python/Copy_File_Control/temp.txt')

# 화일 이름 변경-------------------------------------------------------------------

import os
os.rename('C:/GDS_Python/File_Control/temp.txt','renamed_temp.txt')


# import os
# import datetime
#
# Current_Date = datetime.datetime.today().strftime ('%d-%b-%Y')
# os.rename(r'C:\Users\Ron\Desktop\Test\Products.txt',
#           r'C:\Users\Ron\Desktop\Test\Shipped Products_' + str(Current_Date) + '.txt')

import os
import glob
files = glob.glob('abc*.jpg')
for file in files:
    os.rename(file, '{}.txt'.format(???))

# --------------------------------------------------------------------------------


# # 파일카피
#
# import shutil
# shutil.copy('c:/GDS_Python/File_Control/text.txt','c:/Temp.txt')

# # 화일 화일 읽고 쓰기
#
# with open('Text.txt', 'r+') as text_file:
#     print ('The file content BEFORE writing content:')
#     print("-"*40)
#     print (text_file.read())
#     text_file.write(' and I\'m looking for more')
#     print ('The file content AFTER writing content:')
#     print("-" * 40)
#     text_file.seek(0)
#     print (text_file.read())
#
# # close 파일 잊지말 것 메모리 스피드와 관계 있음
#
# text_file = open('Text.txt', 'w')
# text_file.write('''
#
# and I\'m looking for more file writing exercise ...''')