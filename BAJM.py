#coding:utf-8
import xlrd
data = xlrd.open_workbook("F:\JIMU\JMQST.xlsx")
#data = open_excel("F:\JIMU\JMQST.xlsx")
table = data.sheets()[0]  # 第几个sheet
nrows = table.nrows  # 行数
ncols = table.ncols  # 列数
colnames = table.row_values(1)  # 获取第一行数据
print(colnames)
colnames2 = table.row_values(nrows-1) #获取最后一行数据
print(colnames2)
dictn = dict(zip(colnames,colnames2)) #组成字典
del dictn["收益中"] #删除无效数据
print(dictn)
listn = []
for i,j in  dictn.items():
          i = i.replace(',','')
          j = j.replace(',','')
          m = float(i) - float(j) #第一行数据减去最后一行数据获取差值
          listn.append(m) #差值存入listn中
print(listn)
m = 0
for i in listn:
        m += i #求差值的和
print(m)
print("轻松投减少量为:%f"%(m)) #差值和写入文件中
with open('F:\JIMU\pylog', 'a+') as f:
    f.write('\n')
    f.write(str(m))
print("sucessful ok")