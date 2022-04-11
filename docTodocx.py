# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

from win32com import client as wc
from pydocx import PyDocX      # docx转html用

'''
doc文件转docx文件
fullpath:路径+文件名(不带后缀)
如：D:\\test\\文件1
'''
import os
from win32com import client as wc

word = wc.Dispatch('Word.Application')
# doc文件路径，如C:\\Users
filePath = ""
'''
定位文件夹下所有doc文件
'''
def fileSearch(path):
    filesList = []
    for root, dirs, files in os.walk(path):
        # 若还有子文件夹则遍历
        for dir in dirs:
            fileSearch(path + '\\' +dir)
        # 遍历所有doc文件
        for file in files:
            # 判断尾缀是不是doc
            suffix = file.split('.')[1]
            if suffix == 'doc':
                docToDocx(path + '\\' + file)
'''
将doc文件转换成docx文件,并且删除原来的文件
'''
def docToDocx(filePath):
    try:
        print("开始处理     文件名：" + filePath)
        doc = word.Documents.Open(filePath)
        # [:-4]的意思是选这个字符串从开始到最后倒数第4位（不含）
        docxNamePath = filePath.split('.')[0] + '\\' + filePath[:-4] + '.docx'
        doc.SaveAs(filePath.split('.')[0], 12, False, "", True, "", False, False, False, False)
        # 删除原来的doc文件
        os.remove(filePath)
    finally:
        # 一定要记得关闭docx，否则会出现文件占用
        doc.Close()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    try:
        fileSearch(filePath)
    finally:
        word.Quit()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
