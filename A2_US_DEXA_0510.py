# ......... .... .... .... .... .... .... .... .... .... .... ......
# .... GDS Clinic Routine Check Management Program version 1.30 ....
# ......... .... .... .... .... .... .... .... .... .... .... ......

import pickle
import xlrd
from K_main_mymodules import Fa_pre,Fa_rwb,Fa_val
FS = Fa_rwb.file_save
# --------------------------------------------------------------------------------saving definition
Fa_pre.status_check()
# _________________________________________________________________
a_file = open('C:/GDSRC/data.pkl', "rb")
output = pickle.load(a_file)
a_file.close()

name = (output['R_name'])
b_name = name[:-7]
gender = (output['R_gender'])
date_c = (output['R_date'])
age = (output['R_age'])
# -------------------------------------------------------------------------------extract excel data
file_location =  (Fa_val.br_so_20)
workbook = xlrd.open_workbook(file_location)
# worksheet = workbook.sheet_by_name('Main')
worksheet = workbook.sheet_by_name('Sheet1')

JJ_excel = worksheet._cell_values

sel_name = name
d_sheet_count = worksheet.nrows
for row in (JJ_excel[1:d_sheet_count]):
    keyword = row[0]
    koh_dict = {}
    koh_dict = dict({'R_ID': row[5], 'R_name': row[6], 'R_gender': row[7],'R_doctor': row[9],'R_date':row[10]})
    koh_dict.update({'R_height': row[69], 'R_weight': row[70],'R_sbp': row[74],'R_dbp': row[75]})
    koh_dict.update({'R_CSF_biopsy': row[66], 'R_FVC': row[81], 'R_FEV1': row[83], 'R_FEV1_FVC': row[84]})
    koh_dict.update({'R_dexaJ': row[44],'R_z_scoreJ': row[94],'R_t_score1': row[91],'R_t_score2': row[92]})
    koh_dict.update({'R_t_score3': row[93],'R_t_score4': row[95]})

    Woon = (list(koh_dict.values()))
# ------------------------------------------------------------------------------------USG PREGRAM
    name = (row[6])
    if sel_name == name:
        GFS = row[58]
        CFS = row[66]
        PAP = row[47]

        FS('\n\n' + '-'*60 + 'Chest PA\n' + row[113])
        if gender =='f':
            FS('\n\n' + '-'*60 + 'Mammography\n' + row[114])
            FS('\n\n' + '-'*60 + 'BUS\n' + row[134])
        FS('\n\n' + '-'*60 + 'Thyroid US\n' + row[133])
        FS('\n\n' + '-'*30 + 'US Aorta\n' + row[125])
        FS('\n\n' + '-'*30 + 'CUS carotid artery\n' + row[135])
        FS('\n\n' + '-'*60 + 'Others US\n' + row[128])
        FS('\n\n' + '-'*30 + 'US Liver\n' + row[123])
        FS('\n\n' + '-'*30 + 'US Biliary system\n' + row[124])
        FS('\n\n' + '-'*30 + 'US Pancreas\n' + row[125])
        FS('\n\n' + '-'*30 + 'US Kidney\n' + row[127])
        FS('\n\n' + '-'*60 + 'US conclusion\n' + row[130])
        # --------------------------------------------------------------------------------------GFS
        def biopsy_report(x, y):
            file_location = (Fa_val.br_ex_20 + 'D_2020_' + b_name + '_.xlsx')
            workbook = xlrd.open_workbook(file_location)
            ws_tis = workbook.sheet_by_name(x)
            J_tis = ws_tis._cell_values

            N = len(J_tis)
            for i in range(N):
                J = J_tis.pop(0)
                if y in J[3] and 'C' in J[4]:
                    FS(J[3])

        FS('\n\n' + '-'*60 + 'GFS\n')
        FS('Esophagus  :  ' + row[56]+'\n')
        FS('Stomach    :  ' + GFS + '\n')
        FS('Duodenum   :  ' + row[59] + '\n')
        FS('\nConclusion :  \n' + row[63] + '\n\n')

        if row[61] == False:
            FS('CLO test   :  헬리코박터균 검사는 시행하지 않았습니다.\n')
        else: FS('CLO test   : 헬리코박터균 검사는 ' + str(row[62]) + '입니다.' + '\n')

        if row[60] == False:
            GFS_B = ' 상부위장관 내시경에서 조직검사는 시행하지 않았습니다.'
            FS('Biopsy     : ' + GFS_B + '\n')
        else:
            biopsy_result = []
            FS('\n' + 'Biopsy     : 위내시경에서 조직검사를 시행하였습니다.' + '\n')
            FS('\n' + '__________________________' + '\n')

            biopsy_report('Tissue','Stomach')
            biopsy_report('Tissue','Esophagus')
            biopsy_report('Tissue','Duodenum')

    # ------------------------------------------------------------------------------------------CFS
        if CFS != '':
            FS('\n\n' + '-'*60 + 'CFS\n')
            # row[65]  = biopsy 시행 여부
            FS('대장내시경 결과  : \n' + row[57])
            FS('\n\n결론 :  \n'+ row[66])

            if row[65] == False:
                CFS_B = '대장내시경에서 조직 검사는 시행하지 않았습니다.'
                FS('\n\nBiopsy     : ' + CFS_B + '\n')
            else:
                biopsy_result = []
                b_name = name[:3]
                FS('\n\nBiopsy     : 대장내시경에서 조직검사를 시행하였습니다.' + '\n')
                FS('\n' + '__________________________' + '\n')
                biopsy_report('Tissue', 'Colon')
                biopsy_report('Tissue', 'colon')
                biopsy_report('Tissue', 'intestine')
        # ----------------------------------------------------------------------------------------- PAP
        if PAP == True:
            FS('\n\n' + '-'*60 + 'PAP\n')
            FS('\n' + '< 자궁경부암 검진 >\n')
            biopsy_report('PAP','자궁경부')
    # ---------------------------------------------------------------------------------------- DEXA
        import subprocess
        subprocess.call(['cls'], shell=True)
        if row[91] == '':
            continue
        if sel_name == name and gender == 'f':
            FS('\n\n' + '-' * 60 + 'DEXA\n')

            z_score = (row.pop(94))
            s_row = sorted(row[91:94])

            t_score = s_row[0]
            dexa_J = (row[44])

            if dexa_J == True :
                # premeno = 'n'
                # fracture = 'n'
                while True:
                    print("-"*75 + 'Osteoporosis')
                    premeno = input("당신은 \n\n18세 미만의 소아 혹은 청소년,\n\n폐경기 전의 여성 \n\n혹은 50세 미만의 남성 입니까?\n\n   yes / no  ( y/n )   ")
                    fracture = input("\n당신은 골절 병력이 있습니까?\n\n   yes / no  ( y/n )   ")
                    if premeno in ['y', 'n'] and fracture in ['y', 'n']:
                        break
                if premeno =="y" :
                    if z_score <= -2.0:  # "연령기대치이하"
                        z_score = format(z_score,'5.2f')
                        FS("\n골밀도 검사 결과 Z-score = {0} 입니다." + format(z_score,'5.2f'))  #--save result
                        z_judge = open('C:/GDSRC/Ref_Text/gdsdexJL.txt', 'r', encoding='utf-8')

                        for z_judgeC in z_judge:
                            z_judgeCr = z_judgeC.rstrip()
                            if z_judgeCr.startswith('dexaC01'):
                                Z_Score = ("\n{0} ".format(z_judgeCr[9:]))
                                FS(Z_Score)     #----------------save result
                            else:
                                break
                    else:
                        z_score = format(z_score,'5.2f')
                        FS('\n골밀도 검사 결과 Z-score = {0} 입니다.'.format(z_score))
                        FS('\n정상 입니다.')
                        break
                else:
                    FS("\n골밀도 검사 결과 T-score = {0:5.2f} 입니다.".format(t_score))
                    if fracture == "y":
                        if t_score <= -2.5:
                            dexa_JL = 'dexaC03'
                            break
                    else:
                        if t_score <= -2.5:
                            dexa_JL = 'dexaC04'
                            break
                    if -2.5 <  t_score < -1.0: dexa_JL = 'dexaC05'
                    if -1.0 <= t_score:        dexa_JL = 'dexaC06'

                    t_judge = open('C:/GDSRC/Ref_Text/gdsdexJL.txt', 'r', encoding='utf-8')
                    line = t_judge.readlines()
                    for t_judgeC in line:
                        t_judgeCr = t_judgeC.rstrip()
                        if t_judgeCr.startswith(dexa_JL):
                            # print("{0}".format(t_judgeCr[8:]))
                            T_Score = ("\n {0}\n".format(t_judgeCr[8:]))
                            FS(T_Score)
                            # ----------------save result