# -*- coding: UTF-8 -*-

'''
__author__="zf"
__mtime__ = '2016/11/8/21/38'
__des__: 简单的读取文件
__lastchange__:'2016/11/16'
'''
from __future__ import division
import os
import math
from xlwt import Workbook, Formula
import xlrd
import copy
import types
import time

def is_num(unum):
    try:
        unum + 1
    except TypeError:
        return 0
    else:
        return 1


# 不带颜色的读取
def open_file(content):
    # 打开文件
    global workbook, file_excel
    file_excel = str(content)
    if '.xl' not in file_excel:
        file = (file_excel + '.xls')  # 文件名及中文合理性
        if not os.path.exists(file):  # 判断文件是否存在
            file = (file_excel + '.xlsx')
            if not os.path.exists(file):
                print("文件不存在")
    else:
        file = file_excel
        if not os.path.exists(file):
            print("文件不存在")
    workbook = xlrd.open_workbook(file)
    print('suicce')


def read_allmesg(file_name):
    allmesg = {}
    count_num = {}
    allmesg_dict_key = ['document_num', 'model', 'number', 'compose']
    count_num_key = ['sub_model', 'need_number', 'have_number', 'count_compose']
    open_file(file_name)
    Sheetname = workbook.sheet_names()

    for name in range(len(Sheetname)):

            table = workbook.sheets()[name]
            nrows = table.nrows
            for n in range(nrows):
                a = table.row_values(n)
                flag = 0
                for i in range(len(a)):
                    if is_num(a[i]) == 1:
                        if math.modf(a[i])[0] == 0 or a[i] == 0:  # 获取数字的整数和小数
                            a[i] = int(a[i])  # 将浮点数化成整数
                    if name == 0:
                        if a[i] != '':
                            if flag == 0:
                                allmesg[n] = {}
                            allmesg[n][allmesg_dict_key[flag]] = a[i]
                            flag += 1
                    else:
                        if flag == 0:
                            count_num[n] = {}
                        count_num[n][count_num_key[flag]] = a[i]
                        flag += 1


    return allmesg, count_num


def deal_allmesg(allmesg, count_num):
    for no,allmesg in allmesg.item():
        for allmesg_key, allmesg_value in allmesg.item():
            for
if __name__ == "__main__":
    allmesg, count_num = read_allmesg(3)
