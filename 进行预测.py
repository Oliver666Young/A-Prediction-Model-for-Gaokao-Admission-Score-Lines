import xlrd
file = '四川理科历年录取线.xlsx'

#先做2017

uni=[]
quota=[]
uni_rank=[]
uni_rank_least_position=[]
uni_rank_prediction_least_postion=[]

row = 1130 #大学录取线excel的行数
uni_row = 1007 #txt文件的行数

#write uni[]
with open("2017uni_quota(uni).txt", "r",encoding='utf-8') as u:
    for line in u.readlines():                          #依次读取每行  
        line = line.strip()                             #去掉每行头尾空白  
        uni.append(line)

#write quota[]
with open("2017uni_quota(quota).txt", "r",encoding='utf-8') as q:
    for line in q.readlines():                          #依次读取每行  
        line = line.strip()                             #去掉每行头尾空白  
        quota.append(line)

wb = xlrd.open_workbook(filename=file)#打开文件
sheet1 = wb.sheet_by_index(4)#通过索引获取表格

#uni_rank[]
for i in range(1,row):
    uni_rank.append(sheet1.cell_value(i,0))

#uni_rank_least_position[]
#for i in range(1,2266):
#    uni_rank_least_position.append(int(sheet1.cell_value(i,5)))

#uni_rank_prediction_least_postion=[]
for i in range(0,row-1):
    p = 0
    for n in range(0,i+1):
        for t in range(0,uni_row):
            if uni_rank[n] == uni[t]:
                a = quota[t]
                p = p + int(a)
    uni_rank_prediction_least_postion.append(p)

print(uni_rank_prediction_least_postion[0])
print(uni_rank_prediction_least_postion[1])

with open("2017prediction.txt","w") as t:
    for i in range(0,len(uni_rank_prediction_least_postion)):
        n = str(uni_rank_prediction_least_postion[i])
        t.write(n)
        t.write("\n")