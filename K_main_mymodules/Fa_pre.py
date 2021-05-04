# Februarius

import os, sys
import subprocess
import pandas as pd

def status_check():
    # print('Current Directory   :   ', os.getcwd())
    os.chdir('C:/GDSRC')
    sys.path.append('./K_main__mymodules')
    subprocess.call("cls", shell=True)
    # subprocess.call("dir ", shell=True)

    # print('\nPython Version      :   ', sys.version)
    # print('pandas Version      :   ', pd.__version__)
    # print('Working Directory   :   ', os.getcwd())

def pt(x,y):
    print(x + '\t\t\t\t.............................', type(y))

