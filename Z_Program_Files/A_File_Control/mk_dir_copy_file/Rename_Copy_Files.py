# ------------------------------------------------directory making
import datetime, os
from os import makedirs
name = input(('이름은 ?   '))

today = datetime.date.today()
todaystr = today.isoformat()

mfn_date = todaystr.split('-')
f_date = (mfn_date[0]+'_'+ mfn_date[1])
file_date = (mfn_date[0] + '_' + mfn_date[1] + '_' + mfn_date[2])

print((todaystr))
print(f_date)
print(file_date)

# os.makedirs('C:/GDSRC/Data/' + file_date + '_' + name, exist_ok=True)
# os.makedirs('C:/GDSRC/Result/' + file_date + '_' + name, exist_ok=True)

os.makedirs('c:/GDSRC/Data/' + file_date + '_' + name, exist_ok=True)
os.makedirs('c:/GDSRC/Result/' + file_date + '_' + name, exist_ok=True)

# ---------------------------------------------file rename and copy
import os,glob, datetime
# name = input('이름은 ?   :  ')
Current_Date = datetime.datetime.today().strftime('%Y_%m_%d' + '_' + name)
print(Current_Date)
src = Current_Date
os.chdir('C:/GDSRC/')

# for file in os.listdir():
#     print(file)
#     if file == ('KO'+'J'+'.xls'):
#         dst = 'C:/GDS/REn_Ren_file.xls'
#         os.rename(file,dst)
#         print(dst)

for f in glob.glob(name +"*"):
      # os.chdir('C:/GDSRC/Result/' + src + '/')
      os.rename(f, 'C:/GDSRC/Data/' + src + '/'+ src + ".xls")

