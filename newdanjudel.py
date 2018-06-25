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

def get_allfile_name():
    produce_compose=[]
    number_produce=[]
    name_type = ['MRP报告数据', '生产建议数据']
    pcw = os.getcwd()
    all_name = os.listdir(pcw)
    for x in all_name:
        for y in range(len(name_type)):
            if name_type[y] in x:
                if y == 0:
                    produce_compose = read_produce_compose(x, produce_compose)
                else:
                    number_produce = read_papernumber(x, number_produce)

    return produce_compose, number_produce


def is_num(unum):
    try:
        unum + 1
    except TypeError:
        return 0
    else:
        return 1


# 不带颜色的读取
def filename(content):
    # 打开文件
    global workbook, file_excel
    file_excel = str(content)
    file = (file_excel )  # 文件名及中文合理性
    if not os.path.exists(file):  # 判断文件是否存在
        # file = (file_excel + '.xlsx')
        if not os.path.exists(file):
            print("文件不存在")
    workbook = xlrd.open_workbook(file)
    print('suicce')


def read_produce_compose(content, produce_compose):
    allitem = produce_compose
    filename(content)
    produce_compose= []
    # 获取所有的sheet
    Sheetname = workbook.sheet_names()
    # print "文件",file_excel,"共有",len(Sheetname),"个sheet："

    for name in range(len(Sheetname)):

        table = workbook.sheets()[name]
        nrows = table.nrows
        for n in range(nrows):
            mid = []
            # 获取单行内容
            a = table.row_values(n)
            for i in range(len(a)):

                if is_num(a[i]) == 1:
                    if math.modf(a[i])[0] == 0 or a[i] == 0:  # 获取数字的整数和小数
                        a[i] = int(a[i])  # 将浮点数化成整数

                mid.append(a[i])
            produce_compose.append(mid)
    allitem = allitem + produce_compose
    return allitem


def read_papernumber(content, number_produce):
    allitem = number_produce
    filename(content)
    number_produce = []
    number_produce_list =[]
    # 获取所有的sheet
    Sheetname = workbook.sheet_names()
    # print "文件",file_excel,"共有",len(Sheetname),"个sheet："

    for name in range(len(Sheetname)):

        table = workbook.sheets()[name]
        nrows = table.nrows
        for n in range(nrows):
            mid = []
            # 获取单行内容
            a = table.row_values(n)
            for i in range(len(a)):

                if is_num(a[i]) == 1:
                    if math.modf(a[i])[0] == 0 or a[i] == 0:  # 获取数字的整数和小数
                        a[i] = int(a[i])  # 将浮点数化成整数

                mid.append(a[i])
            number_produce_list.append(mid)


    #  将每行信息转换成一个字典，存放在list danzi里
    titletime = 0
    for x in range(len(number_produce_list)):
        tiaomuxijie = {}
        if titletime == 0:
            if number_produce_list[x][4] != u'':
                print('succeeeee')
                titlehang = number_produce_list[x]
                titletime = 1
        if titletime == 1:
            for y in range(len(number_produce_list[x])):
                tiaomuxijie[titlehang[y]] = number_produce_list[x][y]
            if tiaomuxijie != {}:
                number_produce.append(tiaomuxijie)

    allitem = allitem + number_produce
    return allitem




def read_specal_self_made(content):
    filename(content)

    Sheetname = workbook.sheet_names()
    specal_self_made =[]
    for name in range(len(Sheetname)):

        table = workbook.sheets()[name]
        nrows = table.nrows
        for n in range(nrows):
            mid = []
            # 获取单行内容
            a = table.row_values(n)
            for i in range(len(a)):

                if is_num(a[i]) == 1:
                    if math.modf(a[i])[0] == 0 or a[i] == 0:  # 获取数字的整数和小数
                        a[i] = int(a[i])  # 将浮点数化成整数

                specal_self_made.append(a[i])
    return specal_self_made


# print zizhi_peijian,'bbbbbbbbbbbbbbbbbbbbbb'
def delzizhijian(produce_compose, specal_self_made):
    conpose_produce_mesg=[]
    typename = [u'ZP-4', u'ZP-7', u'ZP-8', u'ZP-9']
    aaa = [u'0', u'1', u'2', u'3', u'4', u'5', u'6', u'7', u'8', u'9']
    mid = []
    for x in range(len(produce_compose)):

        global title_namex
        if len(produce_compose[x]) > 3 and type(produce_compose[x][0]) == type(4):
            if produce_compose[x][1] == '' and produce_compose[x][0] != '' and produce_compose[x][2] != '':
                if mid != []:

                    if u'ZP' not in mid[0]:
                        conpose_produce_mesg.append(mid)

                mid = []
                title_name = produce_compose[x][2]
                mid.append(title_name)
            if type(produce_compose[x][1]) == type(4) and u'ZNP' in title_name and type(produce_compose[x][2]) != type(4):
                panduantitle = produce_compose[x][2][:4]
                if produce_compose[x][1] == 1 and panduantitle in typename and (
                        (len(produce_compose[x][2]) == 7 and produce_compose[x][2][-1] in aaa) or len(produce_compose[x][2]) > 7):
                    if len(mid) == 1:
                        mid.append(produce_compose[x][2][3:])
                        mid.append(produce_compose[x][5])
                    elif len(mid) >= 2:
                        mid[1] = mid[1] + u'/' + produce_compose[x][2][3:]

                elif produce_compose[x][2] in specal_self_made:
                    if len(mid) == 1:
                        mid.append(produce_compose[x][2][3:])
                        mid.append(produce_compose[x][5])
                    elif len(mid) >= 2:
                        mid[1] = mid[1] + u'/' + produce_compose[x][2][3:]

    if u'ZP' not in mid[0]:
        conpose_produce_mesg.append(mid)
    return conpose_produce_mesg


def add_papernumber(conpose_produce_mesg, number_produce):
    for x in conpose_produce_mesg:
        for y in number_produce:
            if x[0] == y[u'存货编号'] and x[2] == y[u'净需求量'] and 'flag' not in y:
                x.append(y[u'已下达单据号'])
                y[u'flag'] = 1
                break
    differnt_compose_item = check_lack(conpose_produce_mesg)
    return conpose_produce_mesg, differnt_compose_item


def check_lack(conpose_produce_mesg):
    differnt_compose_item = {}
    for x in conpose_produce_mesg:
        # if len(x) < 3:
        #     x.append('')
        #     x.append('')

        if u'/' in x[1]:
            need_deal_peijian = x[1].split(u'/')
        else:
            mid = []
            mid.append(x[1])
            need_deal_peijian = mid

        for y in need_deal_peijian:
            if y not in differnt_compose_item:
                differnt_compose_item[y] = x[2]
            else:
                differnt_compose_item[y] = differnt_compose_item[y] + x[2]
    return differnt_compose_item

def chage_place(newoutitem):
    final_zizhijian = []
    for x in newoutitem:
        mid =[]
        mid.append(x[3])
        mid.append('')
        mid.append(x[0])
        mid.append(x[2])
        mid.append('')
        mid.append('')
        mid.append(x[1])
        final_zizhijian.append(mid)
    return final_zizhijian
def out_mesg(output_item, differnt_compose_item):
    a = []
    print(len(a))

    book = Workbook()
    sheet1 = book.add_sheet(u'自制件')
    for i in range(len(output_item)):
        for j in range(len(output_item[i])):

            if is_num(output_item[i][j]) == 1:
                if math.modf(output_item[i][j])[0] == 0 or output_item[i][j] == 0:  # 获取数字的整数和小数
                    output_item[i][j] = int(output_item[i][j])  # 将浮点数化成整数
            sheet1.write(i, j, output_item[i][j])

    sheet2 = book.add_sheet(u'统计')
    i = 0
    for key, value in differnt_compose_item.items():
        sheet2.write(i, 0, key)
        sheet2.write(i, 1, value)
        i = i + 1

    book.save('3.xls')  # 存储excel
    book = xlrd.open_workbook('3.xls')
    print('----------------------------------------------------------------------------------------')
    print('----------------------------------------------------------------------------------------')
    print(u'计算完成')

    print('----------------------------------------------------------------------------------------')

    print('----------------------------------------------------------------------------------------')

    time.sleep(10)


def sort_list(outitem):
    for x in outitem:
        x[0] = x[0].split('ZNP-')[1]
    newoutitem = sorted(outitem, key=lambda x:x[0])
    return newoutitem




if __name__ == "__main__":
    produce_compose, number_produce = get_allfile_name()
    # produce_compose = read_produce_compose('1')
    specal_self_made = read_specal_self_made('自制配件.xls')
    # number_produce = read_papernumber('2')
    conpose_produce_mesg = delzizhijian(produce_compose, specal_self_made)
    conpose_produce_mesg, differnt_compose_item = add_papernumber(conpose_produce_mesg, number_produce)
    newoutitem = sort_list(conpose_produce_mesg)
    output_item = chage_place(newoutitem)
    out_mesg(output_item, differnt_compose_item)
# readnew2('new2')
