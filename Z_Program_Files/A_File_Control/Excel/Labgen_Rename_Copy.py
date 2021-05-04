

# ---------------------------------------------file rename and copy
import os,glob, datetime

name = ('')
Current_Date = datetime.datetime.today().strftime('%Y_%m_%d' + '_' + name)
print(Current_Date)
src = Current_Date
os.chdir('C:/GDSRC/Data/')

for file in os.listdir():
    print(file)
    if file.startswith('지디스'):
          print(file)
          os.rename(file, 'C:/GDSRC/Data/' + src + 'Labgen.xls')

