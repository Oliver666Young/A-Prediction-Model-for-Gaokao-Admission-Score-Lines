import xlrd

##记得修改写的文件！！

file = '2017四川理科招生计划.xlsx'
uni = []
quota = []
num = 2409 #大学数量
row = 22880 #表格行数

for i in range (0,num):#大学数量
    quota.append(0)

#def uni_quota():

wb = xlrd.open_workbook(filename=file)
sheet1 = wb.sheet_by_index(0)

for i in range(1,row):#表格行数
    s = str(sheet1.cell(i,0))
    if uni.count(s[6:len(s)-1]) == 0:
        uni.append(s[6:len(s)-1])
#finish uni[]

for i in range(1,row):
    s = str(sheet1.cell(i,0))
    ns = str(sheet1.cell(i,7))
    if ns[0] == "t":
        n = int(ns[6:len(ns)-1])
    else:
        n = 0
    for j in range (0,num):
        if s[6:len(s)-1] == uni[j]:
            quota[j] = quota[j] + n 
#finish quota

with open("2017uni_quota(uni).txt","w", encoding='utf-8') as f:
    for i in range(0,len(uni)):
        s = str(uni[i])
        f.write(s)
        f.write("\n")

with open("2017uni_quota(quota).txt","w", encoding='utf-8') as f:
    for i in range(0,len(quota)):
        n = str(quota[i])
        f.write(n)
        f.write("\n")