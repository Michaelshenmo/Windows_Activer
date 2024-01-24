from __future__ import print_function
import ctypes, sys
import os
import pandas as pd
import getpass
if ctypes.windll.shell32.IsUserAnAdmin() == 0:
    print('此程序需要管理员权限，请在询问时点击‘是’。')
    print('如果你无法授予此权限，请联系你的系统管理员。')
    a = ''
    disk = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    for i in range(26):
        if 'Active.exe' in os.popen('dir ' + disk[i] + ':').read():
            a = disk[i]
    os.system('start ' + a + ':\\Admin.cmd')
if ctypes.windll.shell32.IsUserAnAdmin() == 1:
    df = pd.read_excel('Windows_Active_Key.xlsx')
    os.system('start winver')
    pd.set_option('display.max_rows', None)
    print(df[["Version"]])
    b = input('我们打开您电脑的系统版本，找到您的系统版本，输入对应的序号: ')
    try:
        print(df.iat[int(b),2])
    except IndexError:
        ex = getpass.getpass(prompt='输入的序号不在范围内！请重启程序。')
    else:
        print('正在激活您的产品，请注意您的任务栏，弹出窗口时请点击‘确定’。')
        os.system('slmgr -ipk ' + df.iat[int(b),2])
        os.system('slmgr /skms win.kms.pub')
        os.system('slmgr /ato')
        ex = getpass.getpass(prompt='我们完成了激活，请检查是否激活完成，如果没有，请检查你的系统版本，然后重新运行此程序。')