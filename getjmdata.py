#导入所需要的相关包
from html.parser import HTMLParser
from urllib import request
from bs4 import  BeautifulSoup
import re
import time
import xlrd,xlwt,os,sys,xlutils
from xlrd import open_workbook
from xlutils.copy import copy
#此函数为获取盒子的标号及剩余量
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
     listm.append(j.split('/')[2])
     url = "https://box.jimu.com" + j
     with request.urlopen(url) as f:
        data = f.read().decode('utf-8').replace(u'\xa9', u'')
     with open(data3, 'w') as f:
        for i in data:
            f.write(i)
     soup = BeautifulSoup(open(data3))

     jr = (soup.find_all(class_="canbid-amount"))
     if jr:
         listj.append(str(jr[0]).split('<span>')[1].split('</span>')[0])
        
     else:
         listj.append("收益中")
   print(listj)
   print(listm)
   dictjm = dict(zip(listm, listj))
   return dictjm
def sendexcl(url,getjm):
    rexcel = open_workbook(url)
    rows = rexcel.sheets()[0].nrows
    cols = rexcel.sheets()[0].ncols
    excel = copy(rexcel)
    table = excel.get_sheet(0)
    j = 0
    for i, m in getjm.items():
        table.write(rows, j, i)
        j += 1
    rows += 1

    excel.save(url)
    j = 0
    for i, m in getjm.items():
        table.write(rows, j, m)
        j += 1
    rows += 1

    excel.save(url)

getjm = getdata("https://box.jimu.com/Venus/List","jmgetlog","jmget1log")
sendexcl("F:\JIMU\JMQST.xlsx",getjm)

   
   
   
   
   