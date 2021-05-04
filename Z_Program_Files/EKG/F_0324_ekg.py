def file_save(x):
    with open('C:/GDSRC/Result/RC_result_file.txt', 'a',encoding='utf-8') as f:
        for item in (x):
            f.writelines(item)
    f.close()


ekgjudge = open('C:/GDSRC/Ref_Text/gdsekgJL_T.txt','r', encoding="utf-8")
lines = ekgjudge.readlines()
for line in lines:
    print(line[4:])

while True:
    i = input (" Select number   :   .. ")
    i = ('{0:0>2}'.format(int(i)))
    ekg_JL = 'ekgC{}'.format(i)
    i = int(i)
    if i == 0:
        add_EKG = input('다른 결과의 심전도 결과를 넣어 주세요...')
        file_save(add_EKG)
        break
    elif i <= 20: break
    else: print('다시 입력해 주세요 ...!!!')

ekgjudge = open("C:/GDSRC/Ref_Text/gdsekgJL.txt", "r", encoding="utf-8")
lines = ekgjudge.readlines()
for line in lines:
    ekg_res = (line[:6])
    if ekg_res == ekg_JL:
        print(line[7:])
ekgjudge.close()

