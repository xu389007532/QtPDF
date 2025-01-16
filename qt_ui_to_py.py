# -*- coding: utf-8 -*-
'''
	ui转换成py的转换工具
'''

import os
import os.path

# UI文件所在的路径
dir = './'


# 列出目录下的所有ui文件
def listUiFile():
    list = []
    files = os.listdir(dir)
    for filename in files:

        ################掃描ui文件夾###############開始
        # if filename.startswith('ui_'):
            # list2=[i for i in os.listdir(dir+filename) if str(i).endswith('.ui')]
            # list.append(dir+filename+"/"+"".join(list2))
        ################掃描ui文件夾###############結束
        ###############掃描ui文件###############開始
        if filename.endswith('.ui'):
            list.append(filename)
            print(filename)
        ###############掃描ui文件###############結束

    # print("list:",list)
    return list


# 把后缀为ui的文件改成后缀为py的文件名
def transPyFile(filename):
    return os.path.splitext(filename)[0] + '.py'


# 调用系统命令把ui转换成py
def runMain():
    list = listUiFile()
    for uifile in list:
        pyfile = transPyFile(uifile)
        cmd = 'pyuic5 -o {pyfile} {uifile}'.format(pyfile=pyfile, uifile=uifile)
        # print(cmd)
        os.system(cmd)
    cmd2 = 'pyrcc5 -o resourcefile_rc.py resourcefile.qrc'
    os.system(cmd2)
###### 程序的主入口
if __name__ == "__main__":
    runMain()
