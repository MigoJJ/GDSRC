# ......... .... .... .... .... .... .... .... .... .... .... ......
# .... GDS Clinic Routine Check Management Program version 1.30 ....
# ......... .... .... .... .... .... .... .... .... .... .... ......

import os
import pickle
import datetime
import xlrd
from K_main_mymodules import Fa_pre,Fa_rwb,Fa_val
# -----------------------------------------------------------------------------------initialization
a_file = open('C:/GDSRC/data.pkl', "rb")
output = pickle.load(a_file)
a_file.close()
real_data_dict = (output)

print('output   :   ',output)
FS = Fa_rwb.file_save
FR = Fa_val.br_re
Fa_pre.status_check()
real_data = []
koh_dict = {}

os.chdir('C:/GDSRC')
if os.path.exists(FR):
    os.remove(FR)
    print('clearing past result files...\n')
# ----------------------------------------------------------------------------access data extration
Fa_rwb.sa_title()
sel_name = output['R_name']
check_date = output['R_date']
P_name = (" 이름       :    " + sel_name[:-7] + ' '*7 + sel_name[-7:])
FS(P_name)

file_location = (Fa_val.br_so_20)
workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_name('Sheet1')

JJ_excel = worksheet._cell_values
m_sheet_count = worksheet.nrows
for row in (JJ_excel[1:m_sheet_count]):  # -------------------------------------row count selection
    keyword = row[0]
    koh_dict.update({'R_ID': row[5], 'R_name': row[6],'R_gender': row[7], 'R_doctor': row[9]})
    koh_dict.update({'R_height': row[69], 'R_weight': row[70],'R_sbp': row[74], 'R_dbp': row[75], 'R_TBW': row[78]})
    koh_dict.update({'R_FVC': row[81], 'R_FEV1': row[83], 'R_FEV1_FVC': row[84]})
    koh_dict.update({'R_dexaJ': row[44], 'R_chestPA': row[113], 'R_mammo': row[114]})
    koh_dict.update({'R_Jud': row[3]})

    name = (koh_dict['R_name'])
    ID = (koh_dict['R_ID'])
    doctor = (koh_dict['R_doctor'])
    Jud = (koh_dict['R_Jud'])
    gender = (koh_dict['R_gender'])

    height = (koh_dict['R_height'])
    weight = (koh_dict['R_weight'])

    sbp = (koh_dict['R_sbp'])
    dbp = (koh_dict['R_dbp'])
    FVC = (koh_dict['R_FVC'])
    FEV1 = (koh_dict['R_FEV1'])
    FEV1_FVC = (koh_dict['R_FEV1_FVC'])
    TBW = (koh_dict['R_TBW'])
    dexaJ = (koh_dict['R_dexaJ'])
    cheatPA = (koh_dict['R_chestPA'])
    mammo = (koh_dict['R_mammo'])

    Woon = (list(koh_dict.values()))
    if Woon[1] == sel_name:
        #_________ _________________________________gender
        if gender == 1: gender = 'm'
        elif gender == 0: gender = 'f'
        # _________________________________________age date
        year     = (datetime.date.today().year)
        ID = str(ID)
        age      = (year - (int(ID[0:2])) - 1900)
        p_gender = ('성별        :       ' + gender + '\n')
        p_age    = ('\n\n나이        :     ' + str(age) + '세\n')

        today = datetime.datetime.today()
        c_dt = today.strftime("%B %d, %Y")

        for i in [
            (p_age),
            (p_gender + '\n'),
            ('ID        :   ' + ID[:-2]),
            ("\n검진일      :      " + check_date),
            ("\n판정날      :   " + c_dt)]:
            FS(i)

        for j in [
            {'R_name': name},
            {'R_gender': gender},
            {'R_age': age},
            {'R_ID': ID},
            {'R_date': check_date}]:
            real_data_dict.update(j)

#  ----------------------------------------------------------------------------------------Pickling
        import pickle
        dictionary_data = real_data_dict
        a_file = open('C:/GDSRC/data.pkl', "wb")
        pickle.dump(dictionary_data, a_file)
        a_file.close()
#  -----------------------------------------------------------------------------------------Summary
        FS('\n\n' + '-'*60 + 'Summary'+'\n')
        FS(Jud)
# ---------------------------------------------------------------------------------------------BMI
        if gender == 'm':
            std_weight = height * height * 20 / 10000
        else:
            std_weight = height * height * 21 / 10000

        bmi = round(weight / (height/100)**2, 1)
        std_weight = round(std_weight, 1)
        red_weight = round(std_weight-weight, 1)
        weight = ("{:5.1f}".format(weight))

        if bmi >= 35:           bmi_judge = "고도 비만"
        elif 35 > bmi >= 30:    bmi_judge = "중정도 비만"
        elif 30 > bmi >= 25:    bmi_judge = "경도 비만"
        elif 25 > bmi >= 23:    bmi_judge = "과체중"
        elif 23 > bmi >= 18.5:  bmi_judge = "정상체중"
        else:                   bmi_judge = "저체중"

        BMI_R = ("\n체질량지수의 판독 결과는 {4} 입니다.\n"
                 "당신의 신장은 {5} cm 이고, 현재 체중은 {2} kg 입니다.\n  따라서 체질량지수..BMI..는 {1} 로 {4}입니다.\n"
                 "표준 체중은 {0} kg 으로 {3:>+4}kg 조절하시는 것이 좋겠습니다.\n "
                 .format(std_weight, bmi, weight, red_weight, bmi_judge, height))

        FS('\n\n' + '-'*60 + 'BMI')
        FS(BMI_R)

#         # TBW = input('\n\n당신의 복부 둘레는 ?    :   ')
        TBW = row[78]
        print(row[78])
        if gender == 'f': std_cm = 85
        if gender == 'm': std_cm = 90
        if int(TBW) >= std_cm: TBW_R = (' 당신은 복부둘레는 {} cm 입니다. 따라서 복부비만이 있습니다.\n'.format(TBW))
        else:                  TBW_R = (' 당신은 복부둘레는 {} cm 입니다. 복부비만은 없습니다.\n'.format(TBW))

        FS('\n' + '-'*60 + '복부비만\n')
        FS(TBW_R)

        real_data.append(bmi)
#         # --------------------------------------------------------------------------------------HTN
        if dbp < 90 and sbp >= 140:
            bp_judge = "수축기 단독 고혈압"
            bp_judgecomm = "  규칙적인 운동, 식이요법과 함께 자주 혈압을 측정하는 것이 좋겠습니다.\n 계속 높으면 약물치료가 필요할 수도 있습니다.\n"
        else:
            if dbp >= 100 or sbp >= 160:
                bp_judge = "고혈압 2 기 "
                bp_judgecomm = "  고혈압이 있습니다.  규칙적인 운동, 식이요법과 함께 자주 혈압을 측정하는 것이 좋겠습니다.\n약물 요법이 필요하므로 내원하여 상담 받으시기 바랍니다.\n"
            elif (100 > dbp >= 90) or (160 >= sbp >= 140):
                bp_judge = "고혈압 1 기 "
                bp_judgecomm = "  규칙적인 운동, 식이요법과 함께 자주 혈압을 \n  측정하는 것이 좋겠습니다. 약물 요법이 필요하므로 내원하여 상담 받으시기 바랍니다.\n"
            elif (90 > dbp >= 80) or (140 > sbp >= 130):
                bp_judge = "고혈압 전단계 "
                bp_judgecomm = "  규칙적인 운동, 식이요법과 함께 자주 혈압을 \n  측정하는 것이 좋겠습니다. 계속 높으면 약물치료가 필요할 수도 있습니다.\n"
            elif (dbp >= 80) or (130 > sbp > 120):
                bp_judge = "주의 혈압 "
                bp_judgecomm = "  혈압이 약간 높습니다. 규칙적인 운동, 식이요법과 \n  함께 자주 혈압을 측정하는 것이 좋겠습니다.\n"
            elif (dbp <= 80) and (sbp <= 120):
                bp_judge = "정상혈압 "
                bp_judgecomm = "."
            else:
                print("")

        HTN_R = ("\n당신의 혈압은 {0} / {1} mmHg {2}입니다.\n{3}".format(sbp, dbp, bp_judge, bp_judgecomm))

        FS('\n' + '-'*60 + 'Hypertension')
        FS(HTN_R)
# --------------------------------------------------------------------------------------EKG
        print("_"*60)
        FS('\n' + '-'*60 + 'EKG\n')

        ekgjudge = open("C:/GDSRC/Ref_Text/gdsekgJL_T.txt", "r", encoding="utf-8")
        lines = ekgjudge.readlines()
        for line in lines:
            print(line[4:], end="")
        print("_"*60)
        while True:
            ekg = input('EKG 판독 번호를 선택해주세요  : ')
            if ekg == '0':
                ekgC00 = input('새로운 EKG 진단을 넣어 주세요 ..:  ')
                FS(ekgC00)
                break
            elif int(ekg) < 21:
                ekg = ekg.zfill(2)
                break
        i = ekg

        print("_"*60)
        ekg_JL = 'ekgC{}'.format(i)
        ekgjudge = open("C:/GDSRC/Ref_Text/gdsekgJL.txt","r", encoding="utf-8")
        lines = ekgjudge.readlines()
        for line in lines:
            ekg_res = (line[:6])
            if ekg_res == ekg_JL:
                print('\n', line[7:])
                EKG_R = (line[8:])
                FS(EKG_R)
        ekgjudge.close()

        # --------------------------------------------------------------------------------------PFT
        if FVC != '':
            FS('\n' + '-'*60 + 'PFT\n')

            real_data.append(FVC)
            real_data.append(FEV1)
            real_data.append(FEV1_FVC)

            FS('노력성 폐활량 \n     (FVC, Forced vital capacity)                      :  ' + str(FVC) + '%\n')
            FS('1초간 노력성 호기량 \n     (FEV1, forced expired volume in one second)       :  ' + str(FEV1) + '%\n')
            FS('비율 \n     (FEV1/FVC, ratio of the FEV1 / FVC of the lungs.) :  ' + str(FEV1_FVC) + '%\n\n')

            if age >= 60:
                FEV1_FVC = FEV1_FVC + 10
            if FEV1_FVC < 70:
                if FVC >= 80:
                    FS(" 당신의 폐기능은 폐쇄성 장애만 존재하는 것으로 의심됩니다."
                            " 호흡기내과 전문의와 상담하십시오.")
                if FVC < 80:
                    FS(" 당신의 폐기능은 제한성 폐장애를 동반 / 또는 낮은 VC를 갖는 폐쇄성 장애 의심됩니다.\
                            - RV, TLC를 측정하여 LV이 감소하였으면 제한성 폐장애 동반하는 것이고\
                            - LV이 감소하지 않았으면 낮은 VC를 갖는 폐쇄성 장애입니다.\
                            - 호흡기내과 전문의와 상담하십시오.")
            if FEV1_FVC >= 70:
                if FVC >= 80 and FEV1 >= 80:
                    FS(" 당신의 폐기능은 정상입니다.")
                elif FVC >= 80 and FEV1 < 80:
                    FS(" 당신의 폐기능은 폐쇄성 장애 의심됩니다."
                            " 호흡기내과 전문의와 상담하십시오.")
                elif FVC < 80:
                    FS(" 당신의 폐기능은 제한성 폐장애 의심됩니다."
                            " 호흡기내과 전문의와 상담하십시오.")
        else:
            break
    # ---------------------------------------------------------------------pickle dictionary saving