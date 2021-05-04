# Main GDSRC Directory Making

import datetime, os ,glob

while True:
    name_list = []
    name_list.append (input('이름은 ?   '))
    if name_list[-1] == '0':
        break

    print(*name_list)

    today = datetime.date.today()
    todaystr = today.isoformat()

    mfn_date = todaystr.split('-')
    f_date = (mfn_date[0]+'_'+ mfn_date[1])
    file_date = (mfn_date[0] + '_' + mfn_date[1] + '_' + mfn_date[2])

    os.makedirs('C:/GDSRC/Data/' + file_date + '_' + str(*name_list), exist_ok=True)
    os.makedirs('C:/GDSRC/Result/' + file_date + '_' + str(*name_list), exist_ok=True)

    dir_data =  (file_date + '_' + str(*name_list))

print(dir_data)

# ---------------------------------------------------file forwarding ro directory
import shutil
shutil.copy2 ('/GDSRC/NJH.xls', '/GDSRC/Data/'+ dir_data + '/' + dir_data +'.xls')