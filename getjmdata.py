#导入所需要的相关包
from html.parser import HTMLParser
from urllib import request
from bs4 import  BeautifulSoup
import re
import time
import xlrd,xlwt,os,sys,xlutils
from xlrd import open_workbook
from xlutils.copy import copy
#此函数为获取网站的标号及剩余量
def getdata(data1,data2,data3):

   with request.urlopen(data1) as f:
			  data = f.read().decode('utf-8').replace(u'\xa9', u'')
   #listp接收标号的url进行拼接
   listp = []
   res_tr = re.findall(r'/Venus/\d+',data)
   for i in res_tr:
            listp.append(i)
   print(listp)
   listj = []
   listm = []
   #listj接收剩余量 listm接收标号
   for j in  listp:
     listm.append(j.split('/')[2])  #获取/标号
     url = "https://box.jimu.com" + j #拼接标号url
     with request.urlopen(url) as f:
        data = f.read().decode('utf-8').replace(u'\xa9', u'') #获取标号页面
     with open(data3, 'w') as f:  #标号页面存入data3
        for i in data:
            f.write(i)
     soup = BeautifulSoup(open(data3)) #使用BeautifulSoup格式化页面html

     jr = (soup.find_all(class_="canbid-amount")) #根据class标签找到剩余量
     if jr:
         listj.append(str(jr[0]).split('<span>')[1].split('</span>')[0]) #切割剩余量获取数据
        
     else:
         listj.append("收益中") #无剩余量显示收益中
   print(listj)
   print(listm)
   dictjm = dict(zip(listm, listj)) #剩余量 标号存入字典dictjm
   return dictjm  #返回
def sendexcl(url,getjm):
    rexcel = open_workbook(url) #打开excel
    rows = rexcel.sheets()[0].nrows #统计行
    cols = rexcel.sheets()[0].ncols #统计列
    excel = copy(rexcel) #复制表
    table = excel.get_sheet(0) #第一张表
    j = 0
    for i, m in getjm.items(): #字典读取数据excel写入标号
        table.write(rows, j, i) #第一张表行写入数据j用于定位行列 i为插入数据
        j += 1 #列后移一位
    rows += 1 #行后移一位

    excel.save(url) #保存数据
    j = 0
    for i, m in getjm.items(): #excel写入剩余量
        table.write(rows, j, m)
        j += 1
    rows += 1

    excel.save(url)

getjm = getdata("https://xxxxxxxxxxxxxx","jmgetlog","jmget1log") #获取剩余量标号字典
sendexcl("F:\JIMU\JMQST.xlsx",getjm) #excel存入字典

   
   
   
   
   