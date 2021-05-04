# --------------------------------------------------------------------------------import pickle data
import pandas as pd
import pickle
import xlrd,xlwt,openpyxl

a_file = open("data.pkl", "rb")
output = pickle.load(a_file)
print(output)
a_file.close()

name = (output['D_name'])
gender = (output['D_gender'])
age = (output['D_age'])
print(name,gender,age)
# -------------------------------------------------------------------file saving & reading definition
def file_save(x):
    with open('C:/GDSRC/Result/RC_result_file.txt','a',encoding='utf-8') as f:
        for item in (x):
            f.writelines(item)
    f.close()
def file_read(y):
    with open('c:/GDSRC/Result/' + 'Result_biopsy_' + y + '.txt','r',encoding='utf-8') as f:
        GB = f.readlines()
        print(GB)
        file_save(GB)
    f.close()
# ---------------------------------------------------------------------------------extract excel data
file_location = 'C:/GDSRC/Ref_Text/Main_2020.xlsx'
workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_name('Sheet1')
JJ_excel = worksheet._cell_values

sel_name = name
d_sheet_count = worksheet.nrows


# -----------------------------------------------Main data reading PROGRAM
for row in (JJ_excel[1:d_sheet_count]):
    keyword = row[0]
    koh_dict = {}
    koh_dict = dict({'R_ID': row[5], 'R_name': row[6], 'R_gender': row[7],'R_doctor': row[9],'R_date':row[10]})
    koh_dict.update({'R_height': row[69], 'R_weight': row[70],'R_sbp': row[74],'R_dbp': row[75]})
    koh_dict.update({'R_FVC': row[81], 'R_FEV1': row[82],'R_FEV1_FVC': row[84]})
    koh_dict.update({'R_dexaJ': row[44],'R_z_scoreJ': row[94],'R_t_score1': row[91],'R_t_score2': row[92]})
    koh_dict.update({'R_dexaJ': row[44],'R_t_score3': row[93],'R_t_score4': row[95]})

    Woon = (list(koh_dict.values()))

    # ---------------------------------------------------------------------------------------USG PREGRAM
    name = (row[6])
    if sel_name == name:
        mammo = row[114]
        print(row[114:120])
        US_breast = row[134]
        US_abdomen = row[123]
        US_conclusion = row[130]
        GFS = row[58]
        CFS = row[57]
      
        file_save('\n------------------------------------------------------chest PA\n\n')
        chestPA = row[113]
        file_save(chestPA)

        if gender =='f':
            file_save('\n------------------------------------------------------Mammography\n')
            file_save(row[114])
            file_save('\n------------------------------------------------------BUS\n')
            file_save(row[134])
        file_save('\n------------------------------------------------------CUS carotid artery\n')
        file_save(row[135])
        file_save('\n------------------------------------------------------Thyroid US\n')
        file_save(row[133])
        file_save('\n------------------------------------------------------Others US\n')
        file_save(row[127])
        file_save('\n------------------------------------------------------US abdomen\n')
        file_save(US_abdomen)
        file_save('\n------------------------------------------------------US conclusion\n')
        file_save(US_conclusion)
        file_save('\n'+row[124])



        # ------------------------GFS
        file_save('\n------------------------------------------------------GFS\n')
        file_save('Esophagus  :  ' + row[56]+'\n')
        file_save('Stomach    :  ' + GFS + '\n')
        file_save('Duodenum   :  ' + row[59] + '\n')
        file_save('Conclusion :  ' + row[63] + '\n\n')

        if row[61] == False:
            row[61] = '시행하지 않았습니다.'
            file_save  ('CLO test   : 헬리코박터균 검사는 ' + str(row[61]) + '\n')
        else: file_save('CLO test   : 헬리코박터균 검사는 ' + str(row[62]) + '입니다.' + '\n')

        if row[60] == False:
            row[60] = '조직 검사는 시행하지 않았습니다.'
            file_save('Biopsy     : ' + str(row[60]) + '\n')
        else:
            biopsy_result = [] 
            b_name = name[:3]
            file_save('\n' + 'Biopsy     : 조직검사를 시행 하였습니다.' + '\n')
            file_read(b_name)
        # ------------------------CFS
        file_save('\n------------------------------------------------------CFS\n')
        file_save(CFS)
        file_save(row[66])
        
        if row[65] == False:
            row[60] = '조직 검사는 시행하지 않았습니다.'
            file_save('Biopsy     : ' + str(row[60]) + '\n')
        else:
            biopsy_result = [] 
            b_name = name[:3]
            file_save('\n\nBiopsy     : 조직검사를 시행 하였습니다.' + '\n')
            file_read(b_name) 
        
            file_save('_ _ '*25 + 'Biopsy Result')

            file_location = 'C:/GDSRC/April_20_Lab.xlsx'  # -------------Main File Directory
            workbook = xlrd.open_workbook(file_location)
            worksheet = workbook.sheet_by_name('Sheet1')
            JJ_excel = worksheet._cell_values

            def file_save_b(x):
                with open('C:/GDSRC/DATA/Result_biopsy_' + sel_name + '.xlsx', 'a',encoding='utf-8') as f:
                    f.writelines(item)
                f.close()

            sheet_count = worksheet.nrows
            for row in (JJ_excel[1:sheet_count]):  #----------------------row count selection
                keyword = row[0]
                koh_dict = {}
                if row[6].startswith('C') & (row[2] == sel_name):
                    print(row[1:6])
                    file_save_b(row[1:3])
                    file_save_b('\n\n')
                    file_save_b(row[5:6])
                    file_save_b('_'*20 + '\n')  

# -----------------------------------------------DEXA PROGRAM
    if sel_name == name:
        print(row[44])
        z_score = (row.pop(94))
        print(z_score)
        s_row = sorted(row[91:95])
        print(s_row)
        print(s_row[0])
        t_score = s_row[0]
        dexa_J = (row[44])

        if dexa_J == True :

            premeno = input("당신은 \n18세 미만의 소아 혹은 청소년,\n폐경기 전의 여성 \n혹은 50세 미만의 남성 입니까?   yes / no  ( y/n )   ")
            fracture = input("당신은 골절 병력이 있습니까?   yes / no  ( y/n )   ")

            if premeno =="y" :
                # z_score = float(input("Z-score  :  "))
                # print('_' * 60)
                if z_score <= -2.0:  # "연령기대치이하"
                    print("\n골밀도 검사 결과 Z-score = {0} 입니다.".format(z_score))
                    file_save('\n------------------------------------------------------DEXA\n\n')
                    file_save("\n골밀도 검사 결과 Z-score = {0} 입니다.".format(z_score))  #--save result
                    z_judge = open('C:/GDSRC/Ref_Text/gdsdexJL.txt', 'r', encoding='utf-8')
                    for z_judgeC in z_judge:
                                z_judgeCr = z_judgeC.rstrip()
                                if z_judgeCr.startswith('dexaC01'):
                                    print("{0} ".format(z_judgeCr[8:]))
                                    Z_Score = ("\n{0} ".format(z_judgeCr[9:]))
                                    file_save(Z_Score)     #----------------save result
                                else:
                                    break
                else:
                    print("골밀도 검사 결과 Z-score = {0} 입니다.".format(z_score))
                    z_judge = "    정상 입니다."
                    print(z_judge)
                    
            else:
                # t_score = float(input("T-score  :  "))
                print("골밀도 검사 결과 T-score = {0} 입니다.".format(t_score))
                file_save('\n------------------------------------------------------DEXA\n')
                file_save("\n골밀도 검사 결과 T-score = {0} 입니다.".format(t_score))
                if fracture == "y" and t_score <= -2.5:
                    dexa_JL = 'dexaC03'
                if fracture == "n" and t_score <= -2.5:
                    dexa_JL = 'dexaC04'
                if fracture == "n" and -2.5 < t_score < -1.0:
                    dexa_JL = 'dexaC05'
                if fracture == "n" and -1.0 <= t_score:
                    dexa_JL = 'dexaC06'

                t_judge = open('C:/GDSRC/Ref_Text/gdsdexJL.txt', 'r', encoding='utf-8')
                line = t_judge.readlines()
                for t_judgeC in line:
                    t_judgeCr = t_judgeC.rstrip()
                    if t_judgeCr.startswith(dexa_JL):
                        print("{0}".format(t_judgeCr[8:]))
                        T_Score = (" {0}\n".format(t_judgeCr[9:]))
                        file_save(T_Score)  # ----------------save result