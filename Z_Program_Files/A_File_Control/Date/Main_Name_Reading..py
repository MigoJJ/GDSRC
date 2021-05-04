# ------------------------------------------------rename labgenomics
import xlrd,xlwt,datetime,os
# import glob
from os import makedirs

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

# ------------------------------------------------directory making
file_location = 'C:/GDSRC/Data/Main.xlsx'
workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_name('Sheet1')
KJJ_excel = worksheet._cell_values

for row in (KJJ_excel[1:10]):
    keyword = row[0]
    R_name = row[6]
    R_name = (R_name[:-7])

    today = datetime.date.today()
    todaystr = today.isoformat()

    mfn_date = todaystr.split('-')
    f_date = (mfn_date[0] + '_' + mfn_date[1])
    file_date = (mfn_date[0] + '_' + mfn_date[1] + '_' + mfn_date[2])
    FR_name = (file_date + '_' + R_name)

    os.makedirs('C:/GDSRC/Data/' + FR_name, exist_ok=True)
    os.makedirs('C:/GDSRC/Result/' + FR_name, exist_ok=True)

    # print(f_date)
    # print(file_date)
    # print(R_name)

# ------------------------------------------------labgenomics parsing

