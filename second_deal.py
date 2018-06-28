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


def old_deal_allmesg(allmesg, count_num):
    for count_num_no in count_num:
        temp_coust_dict ={}
        temp_have = []
        left_num = count_num[count_num_no]['have_number']
        #为所有的信息加上ok这个键
        for x in allmesg:
            allmesg[x].setdefault('ok', '')

        if left_num !=u'':
            for allmesg_no in allmesg:
                if count_num[count_num_no]['sub_model'] in allmesg[allmesg_no]['compose']:
                    temp_coust_dict[allmesg_no] = allmesg[allmesg_no]['number']
            # 将得到的所有组合按照消耗数量进行排序
            sorted_count_item = sorted(temp_coust_dict.items(), key=lambda d: d[1])
            # 从小到大对对应的数据进行处理
            for x in sorted_count_item:
                if allmesg[x[0]]['ok'] != '':
                    temp_have.append('0')
                else:
                    if left_num >= x[1]:
                        left_num -= x[1]
                        temp_have.append('1')
                    else:
                        temp_have.append('0')
            if count_num[count_num_no]['sub_model'] == '8473':
                print (temp_have, sorted_count_item)

            for y in range(len(temp_have)):
                if temp_have[y] == '1':
                    if allmesg[sorted_count_item[y][0]]['ok'] == '':
                        pass
                    else:
                        pass
                elif temp_have[y] == '0':
                    if allmesg[sorted_count_item[y][0]]['ok'] == '':
                        allmesg[sorted_count_item[y][0]]['ok'] = count_num[count_num_no]['sub_model']
                    else:
                        allmesg[sorted_count_item[y][0]]['ok'] = count_num[count_num_no]['sub_model'] \
                                                                 + '/' + allmesg[sorted_count_item[y][0]]['ok']


    return allmesg

def deal_allmesg(allmesg, count_num):
    for x in allmesg:
        allmesg[x].setdefault('ok', '')

    for allmesg_no in allmesg:
        need_deal_conpose = allmesg[allmesg_no]['compose']
        if '/' in need_deal_conpose:
            need_deal_conpose_list = need_deal_conpose.split('/')
        else:
            need_deal_conpose_list = need_deal_conpose
        flag = 0
        for y in need_deal_conpose_list:
            for count_num_no in count_num:
                if y == count_num[count_num_no]['sub_model'] and \
                        isinstance(count_num[count_num_no]['have_number'], int):
                    if count_num[count_num_no]['have_number'] - allmesg[allmesg_no]['number'] >= 0:
                        pass

                    else:
                        flag = 1
                        if allmesg[allmesg_no]['ok'] == '':
                            allmesg[allmesg_no]['ok'] = allmesg[allmesg_no]['ok'] + y
                        else:
                            allmesg[allmesg_no]['ok'] = allmesg[allmesg_no]['ok'] + '/' + y

        if flag == 0:
            for y in need_deal_conpose_list:
                for count_num_no in count_num:
                    if y == count_num[count_num_no]['sub_model'] and \
                            isinstance(count_num[count_num_no]['have_number'], int):
                        count_num[count_num_no]['have_number'] -= allmesg[allmesg_no]['number']

    return allmesg


def output_mesg(allmesg):
    book = Workbook()
    sheet1 = book.add_sheet(u'1')
    i = 0
    num = 1
    for key, value in allmesg.items():
        for s, d in value.items():
            if num == 1:
                sheet1.write(i, 0, key)
            sheet1.write(i, num, d)

            num = num + 2
        num = 1
        i = i + 1

    book.save('4.xls')  # 存储excel
    book = xlrd.open_workbook('4.xls')


if __name__ == "__main__":
    allmesg, count_num = read_allmesg(3)
    allmesg = deal_allmesg(allmesg, count_num)
    output_mesg(allmesg)