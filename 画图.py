import matplotlib.pyplot as plt
import xlrd
file = '新程序\画图.xlsx'

wb = xlrd.open_workbook(filename=file)#打开文件
sheet2017 = wb.sheet_by_index(0)#通过索引获取表格
sheet2018 = wb.sheet_by_index(1)#通过索引获取表格
sheet2019 = wb.sheet_by_index(2)#通过索引获取表格
sheet2020 = wb.sheet_by_index(3)#通过索引获取表格

x_2017 = []
y_2017 = []
for i in range(1,1719):#total#=1718
    x_2017.append(sheet2017.cell_value(i,0))
    y_2017.append(sheet2017.cell_value(i,1))


x_2018 = []
y_2018 = []
for i in range(1,1305):#1720,total=1719
    x_2018.append(sheet2018.cell_value(i,0))
    y_2018.append(sheet2018.cell_value(i,1))

x_2019 = []
y_2019 = []
for i in range(1,1719):#1720,total=1719
    x_2019.append(sheet2019.cell_value(i,0))
    y_2019.append(sheet2019.cell_value(i,1))

x_2020 = []
y_2020 = []
for i in range(1,1719):#1720,total=1719
    x_2020.append(sheet2020.cell_value(i,0))
    y_2020.append(sheet2020.cell_value(i,1))

x_p= []
y_p = []
for i in range(4,15000):#1720,total=1719
    x = 10*i
    x_p.append(x)
    for a in range (1,1729):
        if x < x_2017[a]:
            y1= y_2017[a-1] + (x-x_2017[a-1])*(y_2017[a]-y_2017[a-1])/(x_2017[a]-x_2017[a-1])
            break
    for a in range (1,1305):
        if x < x_2018[a]:
            y2= y_2018[a-1] + (x-x_2018[a-1])*(y_2018[a]-y_2018[a-1])/(x_2018[a]-x_2018[a-1])
            break
    for a in range (1,1729):
        if x < x_2019[a]:
            y3= y_2019[a-1] + (x-x_2019[a-1])*(y_2019[a]-y_2019[a-1])/(x_2019[a]-x_2019[a-1])
            break
    y = (y1+y2+y3)/3
    y_p.append(y)




plt.plot(x_2017,y_2017,color='red')
plt.plot(x_2018,y_2018,color='orange')
plt.plot(x_2019,y_2019,color='blue')
plt.plot(x_2020,y_2020,color='green')
plt.plot(x_p,y_p,color='black')
plt.show()