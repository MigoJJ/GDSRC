def file_save(x):
    with open('C:/GDSRC/Result/RC_result_file.txt','a',encoding='utf-8') as f:
        for item in (x):
            f.writelines(item)
            # f.writelines("%s\n" % item)
    f.close()

import xlrd
file_location = 'C:/GDSRC/Data/Main.xlsx'
workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_name('Sheet1')
JJ_excel = worksheet._cell_values
                                              #-------------patient row selection
sel_name = input('이름을 선택하세요... 이름생일-YY   :   ')
for row in (JJ_excel[1:84]):
    keyword = row[0]
    koh_dict = {}
    koh_dict = dict({'R_ID': row[5], 'R_name': row[6], 'R_gender': row[7],
                    'R_doctor': row[9],'R_date':row[10]})
    koh_dict.update({'R_height': row[69], 'R_weight': row[70],
                    'R_sbp': row[74],'R_dbp': row[75]})
    koh_dict.update({'R_FVC': row[81], 'R_FEV1': row[82],'R_FEV1_FVC': row[84]})
    koh_dict.update({'R_dexaJ': row[44]})

    Woon = (list(koh_dict.values()))
#-------------------------------------DEXA PREGRAM
    koh_dict.update({'R_dexaJ': row[44],'R_z_scoreJ': row[94],'R_t_score1': row[91],'R_t_score2': row[92],
                        'R_t_score3': row[93],'R_t_score4': row[95]})

    name = (row[6])
    if sel_name == name:


        mammo = row[114]
        print(row[114:120])
        US_breast = row[134]
        US_abdomen = row[123]
        US_conclusion = row[130]
        GFS = row[58]
        CFS = row[57]
        # ------------------------chest PA
        file_save('\n------------------------------------------------------chest PA\n\n')
        file_save(row[113])
        print(row[113])
        # ------------------------Mammo
        file_save('\n------------------------------------------------------Mammography\n')
        file_save(row[114])
        # ------------------------BUS
        file_save('\n------------------------------------------------------BUS\n')
        file_save(row[134])
        # ------------------------US
        file_save('\n------------------------------------------------------US abdomen\n')
        file_save(US_abdomen)
        # ------------------------US
        file_save('\n------------------------------------------------------US conclusion\n')
        file_save(US_conclusion)
        # ------------------------GFS
        file_save('\n------------------------------------------------------GFS\n')
        file_save(GFS)
        # ------------------------CFS
        file_save('\n------------------------------------------------------CFS\n')
        file_save(CFS)

    # ----------------------------------------------DEXA

    # if sel_name == name:
        print(row[44])
        z_score = (row.pop(94))
        print(z_score)
        sorted(row[91:95])
        print(sorted(row[91:95]))
        t_score = (row.pop(91))
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
                    break
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
                        break


