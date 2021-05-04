import pandas as pd

path_so  = 'C:/GDSRC_2019/Data/source/'
path_wr  = 'c:/GDSRC_2019/Data/Ianuarius/extraction/2019/'
path_or = 'c:/GDSRC_2019/DATA/source/labgen_2019/Labgen_Data_2019_03.xlsx'

name_list = []
Lab = pd.read_excel(path_or)
#-------------------------------------------------------------------------dupliacate
duplicate_Lab = Lab.drop_duplicates(['수진자명'])
d_lab = duplicate_Lab.iloc[:,2:3]

# print(d_lab)
#-------------------------------------------------------------spliting & list making
cvs_save = d_lab.to_csv(index=False,header=False)
name_list = cvs_save.split('\r\n')
# print(name_list)
# --------------------------------------------------------------------------name insertion
df = pd.read_excel(path_or)
for i in name_list:
    print(i)
    m1 = df[(df['검체명'].isin(['Tissue','Tissue B', 'Tissue C-9', 'Vaginal smear.','Urine (Random)','Stool']) == False) & df['수진자명'].isin([i])]
    m2 = df[df ['검체명'].isin(['Tissue', 'Tissue B', 'Tissue C-9']) & df['수진자명'].isin([i])]
    m3 = df[df ['검체명'].isin(['Vaginal smear.']) & df['수진자명'].isin([i])]
    m4 = df[df ['검체명'].isin(['Urine (Random)']) & df['수진자명'].isin([i])]
    m5 = df[df ['검체명'].isin(['Stool']) & df['수진자명'].isin([i])]

    writer = pd.ExcelWriter(path_wr + 'D_2019_' + i + '.xlsx', engine='xlsxwriter')
    m1.to_excel(writer, sheet_name='Laboratory')
    m2.to_excel(writer, sheet_name='Tissue')
    m3.to_excel(writer, sheet_name='PAP')
    m4.to_excel(writer, sheet_name='Urine')
    m5.to_excel(writer, sheet_name='Stool')
    writer.save()
#----------------------------------------------------------------------------sorting
# sorted_Lab = Lab.sort_values(by=['수진자명','검체명'],ascending=True)
# m = m.sort_values(by=['수진자명','검체명'], axis=0)