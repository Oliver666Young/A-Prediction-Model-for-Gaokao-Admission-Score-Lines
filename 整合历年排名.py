import xlrd
file = '大学排名.xlsx'

wb = xlrd.open_workbook(filename=file)
sheet_rank = wb.sheet_by_index(7) #quota
sheet_2020 = wb.sheet_by_index(11) #quota

#sheet_total = wb.sheet_by_index(5) #总表
#remain_year = []

for i in range (1,1719):#1227
    for k in range (1,597):
        if sheet_rank.cell_value(i,4)<sheet_2020.cell_value(k,2):
            print(sheet_2020.cell_value(k,0))
            break

for i in range (1,1719):#1227
    x3 = sheet_rank.cell_value(i,3)
    # print(x3)
    # print("i is",i)
    for n in range (1,1720):
        if (x3<sheet_2017.cell_value(n,0)):
            x1 = sheet_2017.cell_value(n-1,0)
            y1 = sheet_2017.cell_value(n-1,1)
            x2 = sheet_2017.cell_value(n,0)
            y2 = sheet_2017.cell_value(n,1)
            if((x2-x1)==0):y31 = (y1+y2)/2
            else:y31 = y1 + (x3-x1)*(y2-y1)/(x2-x1)
            #print("y31 is",y31)
            break

    for m in range (1,1720):
        if (x3<sheet_2019.cell_value(m,0)):
            x1 = sheet_2019.cell_value(m-1,0)
            y1 = sheet_2019.cell_value(m-1,1)
            x2 = sheet_2019.cell_value(m,0)
            y2 = sheet_2019.cell_value(m,1)
            if((x2-x1)==0):y31 = (y1+y2)/2
            else:y32 = y1 + (x3-x1)*(y2-y1)/(x2-x1)
            #print("y32 is",y32)
            break
    y3 = (y31+y32)/2
    print(y3)



for i in range (1,2459):#year行数
    remain_year.append(i)

# print(sheet_year.cell_value(445,7))

#历年录取线

for k in range (1,2589):#total行数
    a = 0
    #print(k)
    for i in range (1,2459):
        if sheet_year.cell_value(i,0)==sheet_total.cell_value(k,0) and sheet_year.cell_value(i,4)==sheet_total.cell_value(k,1) and sheet_year.cell_value(i,5)==sheet_total.cell_value(k,2):
            #print(sheet_year.cell_value(i,7))
            #print(i)
            remain_year.remove(i)
            a = 1
            break
    if a == 0:
        print(0)

print(remain_year)        

l = len(remain_year)
for m in range (0,l):
    print(sheet_year.cell_value(remain_year[m],0),sheet_year.cell_value(remain_year[m],4),sheet_year.cell_value(remain_year[m],5),sheet_year.cell_value(remain_year[m],7))


#历年计划数量

b=0
for k in range (1,2890):#total行数
    a=0
    b=b+1
    for i in range (1,2706):#计划行数
        if sheet_year.cell_value(i,0)==sheet_total.cell_value(k,0):
            print(sheet_year.cell_value(i,1))
            a=1
    if a == 0:
        print(0)

print("b is ")
print(b)