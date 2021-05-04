# -----------------------------https://buttercoconut.xyz/265/
# import xlrd,xlwt
#
# file_location = 'C:/GDSRC/EXCEL/20200403.xls'
# workbook = xlrd.open_workbook(file_location)
# worksheet = workbook.sheet_by_name('Sheet1')
# KJJ_excel = worksheet._cell_values
#
# for row in (KJJ_excel[1:300]):
#     keyword = row[0]
#     if row[2] == '오갑선':
#         print(row[2:11])

# -*- coding: utf-8 -*-
import xlrd,xlwt

# 읽기 라이브러리를 통해 현재 같은 디렉터리상에 있는 test.xls를 불러와 workbook에 할당합니다.
workbook = xlrd.open_workbook('2020_04_08_Labgen.xls')

# 워크북에 할당된 엑셀 데이터의 첫번째시트를 불러옵니다.
worksheet = workbook.sheet_by_index(0)

# nrows에 불러온 첫번째 시트의 행수를 불러옵니다.
nrows = worksheet.nrows

# 불러온 데이터를 저장할 딕셔너리를 선언합니다.
datadict = {}

# 행수만큼의 for loop 를 돌려서 행단위로 데이터를 불러와 datadict에 저장합니다.
for row_num in range(nrows):
    datadict[row_num] = {}
    # field_cnt는 열의 개수입니다. 열갯수는 여러분이 지정하시면 됩니다.
    for col in range(11):
        # datadict[row_num]에 열숫자별로 셀데이터를 저장합니다.
        datadict[row_num][col] = worksheet.cell_value(row_num, col)

workbookw = xlwt.Workbook(encoding='utf-8')  # utf-8 인코딩 방식의 workbook 생성

workbookw.default_style.font.height = 20 * 11
worksheetw = workbookw.add_sheet('Sheet10')  # 시트 생성

# 2번째 행부터 데이터를 입력합니다.
for row_num in range(nrows):
    for col in range(11):
        worksheetw.write(row_num, col, datadict[row_num][col])

# 할당된 데이터셋을 'result.xls'로 저장합니다.
workbookw.save("result.xls")