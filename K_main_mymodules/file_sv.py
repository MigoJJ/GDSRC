path_so = './Data/source/'
path_ma = './Data/source/Main-20_June.xlsx'
path_wr = './Data/Ianuarius/extraction/2020/'
path_re = './Result/RC_result_file.txt'

def file_save(x):
    with open('C:/GDSRC/Result/RC_result_file.txt', 'a', encoding='utf-8') as f:
        for item in (x):
            f.writelines(item)
    f.close()

def pr_title():
    file_save("*********************************************************\n\n")
    file_save("               GDS Clinic Routine Check Result           \n\n")
    file_save("*********************************************************\n\n")


def pr_ID():
    file_save(p_age)
    file_save(p_gender + '\n')
    file_save('ID        :    ' + koh_dict['R_ID'])
    file_save("\n검진일      :      " + check_date)
    file_save("\n판정날      :   " + c_dt)
