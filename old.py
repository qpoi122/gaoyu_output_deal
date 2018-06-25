# -*- coding: UTF-8 -*-

'''
__author__="zf"
__mtime__ = '2016/11/8/21/38'
__des__: 简单的读取文件
__lastchange__:'2016/11/16'
'''
from __future__ import division
import xlrd
import os
import math
from xlwt import Workbook, Formula
import xlrd
import sys
import types
import time

def is_chinese(uchar): 
        """判断一个unicode是否是汉字"""
        if uchar >= u'/u4e00' and uchar<=u'/u9fa5':
                return True
        else:
                return False

                
def is_num(unum):
	try:
		unum+1
	except TypeError:
		return 0
	else:
		return 1

#不带颜色的读取
def filename(content):
	#打开文件
	global workbook,file_excel
	file_excel=str(content)
	file=(file_excel+'.xls').decode('utf-8')#文件名及中文合理性
	if not os.path.exists(file):#判断文件是否存在
		file=(file_excel+'.xlsx').decode('utf-8')
		if not os.path.exists(file):
			print "文件不存在"
	workbook = xlrd.open_workbook(file)
	print 'suicce'



def read_produce_compose(content):
	
	filename(content)

		#获取所有的sheet
	Sheetname=workbook.sheet_names()
	# print "文件",file_excel,"共有",len(Sheetname),"个sheet："

	for name in range(len(Sheetname)):

		table = workbook.sheets()[name]
		nrows=table.nrows
		for n in range(nrows):
			mid=[]	
		#获取单行内容
			a=table.row_values(n)
			for i in range(len(a)):	
						
				if is_chinese(a[i]):
					a[i].encode('utf-8' )
				elif is_num(a[i])==1:
					if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
						a[i]=int(a[i])#将浮点数化成整数
				
				mid.append(a[i])
			zizhijian.append(mid)
	# print zizhijian,'bbbbbbbbbbbbbbbbbbbbbb'


def read_papernumber(content):
	danzilist=[]
	filename(content)

		#获取所有的sheet
	Sheetname=workbook.sheet_names()
	# print "文件",file_excel,"共有",len(Sheetname),"个sheet："

	for name in range(len(Sheetname)):

		table = workbook.sheets()[name]
		nrows=table.nrows
		for n in range(nrows):
			mid=[]	
		#获取单行内容
			a=table.row_values(n)
			for i in range(len(a)):	
						
				if is_chinese(a[i]):
					a[i].encode('utf-8' )
				elif is_num(a[i])==1:
					if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
						a[i]=int(a[i])#将浮点数化成整数
					
				mid.append(a[i])
			danzilist.append(mid)




	# print danzilist,'danzidanzidanzidanzidanzidanzidanzi'


	#将每行信息转换成一个字典，存放在list danzi里
	titletime=0
	for x in range(len(danzilist)):
		tiaomuxijie={}
		if  titletime==0:
			if danzilist[x][4]!=u'':
				print 'succeeeee'
				titlehang=danzilist[x]
				titletime=1
		if titletime==1:
			for y in range(len(danzilist[x])):
				tiaomuxijie[titlehang[y]]=danzilist[x][y]
			if tiaomuxijie!={}:
				danzi.append(tiaomuxijie)

	# print danzi,'dnmadnadnandnandnandnand'







def read_specal_self_made(content):
	
	filename(content)

		#获取所有的sheet
	Sheetname=workbook.sheet_names()
	# print "文件",file_excel,"共有",len(Sheetname),"个sheet："

	for name in range(len(Sheetname)):

		table = workbook.sheets()[name]
		nrows=table.nrows
		for n in range(nrows):
			mid=[]	
		#获取单行内容
			a=table.row_values(n)
			for i in range(len(a)):	
						
				if is_chinese(a[i]):
					a[i].encode('utf-8' )
				elif is_num(a[i])==1:
					if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
						a[i]=int(a[i])#将浮点数化成整数
				
				
				zizhi_peijian.append(a[i])
	# print zizhi_peijian,'bbbbbbbbbbbbbbbbbbbbbb'
def delzizhijian():

	typename=[u'ZP-4',u'ZP-7',u'ZP-8',u'ZP-9']
	aaa=[u'0',u'1',u'2',u'3',u'4',u'5',u'6',u'7',u'8',u'9']
	mid=[]
	for x in range(len(zizhijian)):
	
		global title_name
		if len(zizhijian[x])>3 and type(zizhijian[x][0])==type(4):
			if zizhijian[x][1]=='' and zizhijian[x][0]!=''  and zizhijian[x][2]!='' and U'ZNP' in zizhijian[x][2]:
				if mid!=[]:
					# try:
					# 	mid[0]=int(mid[0][4:])
					# except:
					# 	mid[0]=mid[0][4:]
					outitem.append(mid)
				mid=[]
				title_name=zizhijian[x][2]
				mid.append(title_name)
				mid.append(zizhijian[x][5])
			if type(zizhijian[x][1])==type(4) and u'ZNP' in title_name and type(zizhijian[x][2])!=type(4) and u'ZP' in zizhijian[x][2]: 
				#获得判断的本行的产品的头是不是属于自制件的型号里的
				panduantitle=zizhijian[x][2][:4]
				# print len(mid),'llllllllllllllllllllllllllllllllllllllllenenenene'
				if zizhijian[x][1]==1 and panduantitle in typename and ((len(zizhijian[x][2])==7 and zizhijian[x][2][-1] in aaa)  or len(zizhijian[x][2])>7):
					# print len(mid),'llllllllllllllllllllllllllllllllllllllllenenenene'
					if len(mid)==2:
				
						mid.append(zizhijian[x][2][3:])

					elif len(mid)==3:
						# print type(mid[2]),mid[2],'11111',mid[0],zizhijian[x][2][3:],type(zizhijian[x][2][3:])
						mid[2]=mid[2]+u'/'+(zizhijian[x][2][3:])
				elif zizhijian[x][2] in zizhi_peijian:
					if len(mid)==2:
						mid.append(zizhijian[x][2][3:])
					elif len(mid)==3:
						# print type(mid[2]),mid[2],'22111',mid[0],
						mid[2]=mid[2]+u'/'+zizhijian[x][2][3:]

		#全部跑完加上最后一个
		if mid!=[] and x==len(zizhijian)-1:
			# try:
			# 	mid[0]=int(mid[0][4:])
			# except:
			# 	mid[0]=mid[0][4:]
			outitem.append(mid)


	# print outitem,'ooooo'

	


def add_papernumber():
	for x in outitem:
		for y in danzi:
			if x[0]==y[u'存货编号'] and x[1]==y[u'净需求量'] and not y.has_key('flag'):
				x.append(y[u'已下达单据号'])
				y[u'flag']=1
				break
	



def check_lack():
	zizhijian_zucheng={}
	for x in outitem:
		if u'/' in x[2]:
			need_deal_peijian=x[2].split(u'/')
		else:
			mid=[]
			mid.append(x[2])
			need_deal_peijian=mid
		for y in need_deal_peijian:
			if not zizhijian_zucheng.has_key(y):
				zizhijian_zucheng[y]=x[1]
			else:
				zizhijian_zucheng[y]=zizhijian_zucheng[y]+x[1]
	print zizhijian_zucheng

	a=[]
	print len(a)

	book = Workbook()
	sheet1 = book.add_sheet(u'自制件')
	for i in range(len(outitem)):
		for j in range (len(outitem[i])):
			if is_chinese(outitem[i][j]):
				outitem[i][j].encode('utf-8')
			# elif not outitem[i] and outitem[i]!=0: 
			# 	print "空值",
			elif is_num(outitem[i][j])==1:
				if math.modf(outitem[i][j])[0]==0 or outitem[i][j]==0:#获取数字的整数和小数
					outitem[i][j]=int(outitem[i][j])#将浮点数化成整数
			sheet1.write(i,j,outitem[i][j])


	sheet2 = book.add_sheet(u'统计')
	i=0
	for key,value in zizhijian_zucheng.items():

		sheet2.write(i,0,key)
		sheet2.write(i,1,value)
		i=i+1









	book.save('3.xls')#存储excel
	book = xlrd.open_workbook('3.xls')	
	print '----------------------------------------------------------------------------------------'
	print '----------------------------------------------------------------------------------------'
	print u'计算完成'

	print '----------------------------------------------------------------------------------------'

	
	print '----------------------------------------------------------------------------------------'



	time.sleep(10)


if __name__ == "__main__":
	global allitem,newitem,outitem,zizhijian,zizhi_peijian,danzi
	allitem=[]
	newitem=[]
	outitem=[]
	zizhijian=[]
	zizhi_peijian=[]
	danzi=[]
	read_produce_compose('1')
	read_specal_self_made('自制配件')
	read_papernumber('2')
	delzizhijian()
	add_papernumber()
	check_lack()
	# readnew2('new2')

		