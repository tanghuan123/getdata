#coding:utf-8
import xlrd
data = xlrd.open_workbook("F:\JIMU\JMQST.xlsx")
#data = open_excel("F:\JIMU\JMQST.xlsx")
table = data.sheets()[0]  # 第几个sheet
nrows = table.nrows  # 行数
ncols = table.ncols  # 列数
colnames = table.row_values(1)  # 某一行数据
print(colnames)
colnames2 = table.row_values(nrows-1)
print(colnames2)
dictn = dict(zip(colnames,colnames2))
del dictn["收益中"]
print(dictn)
listn = []
for i,j in  dictn.items():
          i = i.replace(',','')
          j = j.replace(',','')
          m = float(i) - float(j)
          listn.append(m)
print(listn)
m = 0
for i in listn:
        m += i
print(m)
print("轻松投减少量为:%f"%(m))
with open('F:\JIMU\pylog', 'a+') as f:
    f.write('\n')
    f.write(str(m))
print("sucessful ok")