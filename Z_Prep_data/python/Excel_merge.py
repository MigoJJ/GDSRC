import pandas as pd

df_L = pd.read_excel('C:/Excel_Merge/LAB.xlsx')
df_B = pd.read_excel('C:/Excel_Merge/BIO.xlsx')
df_P = pd.read_excel('C:/Excel_Merge/PAP.xlsx')

# df.rename(columns = {'old_nm' : 'new_nm'), inplace = True)
df_L.rename(columns = {'결과': '검사결과'},inplace=True)

print(df_L.columns)
df_L_list = df_L.columns.values.tolist()
df_B_list = df_B.columns.values.tolist()
df_P_list = df_P.columns.values.tolist()
print('Laboratory          :   ',df_L_list)
print('Laboratory Length   ;   ',len(df_L))
print('Biopsy              :   ',df_B_list)
print('PAP                 :   ',df_P_list)

df_L = df_L.drop(df_L[df_L['검체명'].isin(['Tissue', 'Tissue B', 'Tissue C-9', 'Vaginal smear.'])].index)
print(len(df_L))
df_L = df_L.drop(['주민번호'],axis=1)
print('Laboratory          :   ',df_L.columns)
print('df_L.shape   :   ',df_L.shape)

frames = [df_L, df_B, df_P]

result = pd.concat(frames)
print(result)
print(type(result))
print(result.shape)
print('Columns   :   ',result.columns)

result.to_excel('C:/Excel_Merge/L2019.xlsx')