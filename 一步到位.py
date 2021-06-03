# coding=UTF-8
import matplotlib.pyplot as plt
import xlrd
import xlwt

excel = xlrd.open_workbook('浙江\数据总表.xlsx')
quota2017 = excel.sheets()[0]
quota2018 = excel.sheets()[1]
quota2019 = excel.sheets()[2]
quota2020 = excel.sheets()[3]
score_line = excel.sheets()[4]
c_table2017 = excel.sheets()[5]
c_table2018 = excel.sheets()[6]
c_table2019 = excel.sheets()[7]
c_table2020 = excel.sheets()[8]
uni_set = []
uni_set_name = []

class Uni:
    uni_name = "未命名"
    uni_ranking_2017 = 1
    uni_ranking_2018 = 1
    uni_ranking_2019 = 1
    uni_ranking_2020 = 1
    uni_ranking_2020p = 0
    quota_2017 = 0#
    quota_2018 = 0#
    quota_2019 = 0#
    quota_2020 = 0#
    score_line_2017 = 0#
    score_line_2018 = 0#
    score_line_2019 = 0#
    score_line_2020 = 0#
    score_line_2020p = 0
    score_ranking_2017 = 0#
    score_ranking_2018 = 0#
    score_ranking_2019 = 0#
    score_ranking_2020 = 0#
    score_ranking_2020p = 0
    # tier_2017 = 0 # 3 denotes having both 1 and 2
    # tier_2018 = 0
    # tier_2019 = 0
    # tier_2020 = 0
    cq_2017 = 0
    cq_2018 = 0
    cq_2019 = 0
    cq_2020 = 0
    cq_2020p =0

# get uni_set
for i in range(1,score_line.nrows):# total 1116 uni
    uni_name = score_line.cell(i,0).value + score_line.cell(i,4).value
    if uni_set.count(uni_name) == 0:
        uni_set.append(uni_name)

# remember name
for i in range(0,len(uni_set)):
    uni_set_name.append(uni_set[i])

# create class
for i in range(0,len(uni_set)):
    uni_set[i] = Uni()
    uni_set[i].uni_name = uni_set_name[i]

    # give value to score line and score ranking
    for k in range(1,score_line.nrows):
        if score_line.cell(k,0).value+score_line.cell(k,4).value==uni_set[i].uni_name and int(score_line.cell(k,3).value)==2017:
            # if uni_set[i].score_line_2017!=0:
            #     a = uni_set_name[i]
            #     a = Uni()
            #     uni_set.append(a)
            #     a.uni_name = uni_set_name[i]
            #     a.score_line_2017 = int(score_line.cell(k,6).value)
            #     a.score_ranking_2017 = int(score_line.cell(k,7).value)
            #     if score_line.cell(k,4).value == "平行录取一段": a.tier_2017 = 1
            #     if score_line.cell(k,4).value == "平行录取二段": uni_set[i].tier_2017 = 2
            # else:
            uni_set[i].score_line_2017 = int(score_line.cell(k,6).value)
            uni_set[i].score_ranking_2017 = int(score_line.cell(k,7).value)
            if score_line.cell(k,4).value == "平行录取一段": uni_set[i].tier_2017 = 1
            if score_line.cell(k,4).value == "平行录取二段": uni_set[i].tier_2017 = 2
            # break
    for k in range(1,score_line.nrows):
        if score_line.cell(k,0).value+score_line.cell(k,4).value==uni_set[i].uni_name and int(score_line.cell(k,3).value)==2018:
            uni_set[i].score_line_2018 = int(score_line.cell(k,6).value)
            uni_set[i].score_ranking_2018 =int(score_line.cell(k,7).value)
            if score_line.cell(k,4).value == "平行录取一段": uni_set[i].tier_2018 = 1
            if score_line.cell(k,4).value == "平行录取二段": uni_set[i].tier_2018 = 2
            # break
    for k in range(1,score_line.nrows):
        if score_line.cell(k,0).value+score_line.cell(k,4).value==uni_set[i].uni_name and int(score_line.cell(k,3).value)==2019:
            uni_set[i].score_line_2019 = int(score_line.cell(k,6).value)
            uni_set[i].score_ranking_2019 = int(score_line.cell(k,7).value)
            if score_line.cell(k,4).value == "平行录取一段": uni_set[i].tier_2019 = 1
            if score_line.cell(k,4).value == "平行录取二段": uni_set[i].tier_2019 = 2
            # break    
    for k in range(1,score_line.nrows):
        if score_line.cell(k,0).value+score_line.cell(k,4).value==uni_set[i].uni_name and int(score_line.cell(k,3).value)==2020:
            uni_set[i].score_line_2020 = int(score_line.cell(k,6).value)
            uni_set[i].score_ranking_2020 = int(score_line.cell(k,7).value)
            if score_line.cell(k,4).value == "平行录取一段": uni_set[i].tier_2020 = 1
            if score_line.cell(k,4).value == "平行录取二段": uni_set[i].tier_2020 = 2
            # break

    # give value to quota
    for k in range(1,quota2017.nrows):
        if quota2017.cell(k,0).value+"平行录取一段"==uni_set[i].uni_name or quota2017.cell(k,0).value+"平行录取二段"==uni_set[i].uni_name:
            uni_set[i].quota_2017 = uni_set[i].quota_2017 + int(quota2017.cell(k,7).value)
    for k in range(1,quota2018.nrows):
        if quota2018.cell(k,0).value+"平行录取一段"==uni_set[i].uni_name or quota2018.cell(k,0).value+"平行录取二段"==uni_set[i].uni_name:
            uni_set[i].quota_2018 = uni_set[i].quota_2018 + int(quota2018.cell(k,7).value)    
    for k in range(1,quota2019.nrows):
        if quota2019.cell(k,0).value+"平行录取一段"==uni_set[i].uni_name or quota2019.cell(k,0).value+"平行录取二段"==uni_set[i].uni_name:
            uni_set[i].quota_2019 = uni_set[i].quota_2019 + int(quota2019.cell(k,7).value)
    for k in range(1,quota2020.nrows):
        if quota2020.cell(k,0).value+"平行录取一段"==uni_set[i].uni_name or quota2020.cell(k,0).value+"平行录取二段"==uni_set[i].uni_name:
            uni_set[i].quota_2020 = uni_set[i].quota_2020 + int(quota2020.cell(k,7).value)

# delete the uni that lacks data in some year
new_uni_set = []
for i in range(0,len(uni_set)):
    if uni_set[i].score_line_2020!=0 and uni_set[i].score_line_2019!=0 and uni_set[i].score_line_2018!=0 and uni_set[i].score_line_2017!=0:
        new_uni_set.append(uni_set[i])

# if both tier 1 and 2, divide the quota
for i in range(0,len(new_uni_set)):
    for k in range(0,len(new_uni_set)):
        a = new_uni_set[i].uni_name
        b = new_uni_set[k].uni_name
        if a[0:len(b)-6]==b[0:len(b)-6] and a[len(b)-6:len(b)]=="平行录取一段" and b[len(b)-6:len(b)]=="平行录取二段" :
            new_uni_set[i].quota_2017 /=2
            new_uni_set[k].quota_2017 /=2
            new_uni_set[i].quota_2018 /=2
            new_uni_set[k].quota_2018 /=2
            new_uni_set[i].quota_2019 /=2
            new_uni_set[k].quota_2019 /=2
            new_uni_set[i].quota_2020 /=2
            new_uni_set[k].quota_2020 /=2

# give value for ranking
for i in range(0,len(new_uni_set)):
    for k in range(0,len(new_uni_set)):
        if new_uni_set[k].score_line_2017 > new_uni_set[i].score_line_2017:
            new_uni_set[i].uni_ranking_2017 = new_uni_set[i].uni_ranking_2017 + 1
        if new_uni_set[k].score_line_2018 > new_uni_set[i].score_line_2018:
            new_uni_set[i].uni_ranking_2018 = new_uni_set[i].uni_ranking_2018 + 1
        if new_uni_set[k].score_line_2019 > new_uni_set[i].score_line_2019:
            new_uni_set[i].uni_ranking_2019 = new_uni_set[i].uni_ranking_2019 + 1
        if new_uni_set[k].score_line_2020 > new_uni_set[i].score_line_2020:
            new_uni_set[i].uni_ranking_2020 = new_uni_set[i].uni_ranking_2020 + 1    


# 2017
cq_2017 = []
ranking_2017 = []
for i in range(0,len(new_uni_set)):
    cq = 0
    for k in range(0,len(new_uni_set)):
        if new_uni_set[k].uni_ranking_2017<=new_uni_set[i].uni_ranking_2017:
            cq = cq + new_uni_set[k].quota_2017
    new_uni_set[i].cq_2017 = cq
    cq_2017.append(cq)
    ranking_2017.append(new_uni_set[i].score_ranking_2017)

cq_2018 = []
ranking_2018 = []
for i in range(0,len(new_uni_set)):
    cq = 0
    for k in range(0,len(new_uni_set)):
        if new_uni_set[k].uni_ranking_2018<=new_uni_set[i].uni_ranking_2018:
            cq = cq + new_uni_set[k].quota_2018
    new_uni_set[i].cq_2018 = cq
    cq_2018.append(cq)
    ranking_2018.append(new_uni_set[i].score_ranking_2018)

cq_2019= []
ranking_2019 = []
for i in range(0,len(new_uni_set)):
    cq = 0
    for k in range(0,len(new_uni_set)):
        if new_uni_set[k].uni_ranking_2019<=new_uni_set[i].uni_ranking_2019:
            cq = cq + new_uni_set[k].quota_2019
    new_uni_set[i].cq_2019 = cq
    cq_2019.append(cq)
    ranking_2019.append(new_uni_set[i].score_ranking_2019)

cq_2020 = []
ranking_2020 = []
for i in range(0,len(new_uni_set)):
    cq = 0
    for k in range(0,len(new_uni_set)):
        if new_uni_set[k].uni_ranking_2020<=new_uni_set[i].uni_ranking_2020:
            cq = cq + new_uni_set[k].quota_2020
    new_uni_set[i].cq_2020 = cq
    cq_2020.append(cq)
    ranking_2020.append(new_uni_set[i].score_ranking_2020)

for i in range(0,len(new_uni_set)):
    new_uni_set[i].uni_ranking_2020p = 0.2*new_uni_set[i].uni_ranking_2017 + 0.52*new_uni_set[i].uni_ranking_2018 + 0.43*new_uni_set[i].uni_ranking_2019

# cq_2020p
for i in range(0,len(new_uni_set)):
    cq = 0
    for k in range(0,len(new_uni_set)):
        if new_uni_set[k].uni_ranking_2020p <= new_uni_set[i].uni_ranking_2020p:
            cq = cq + new_uni_set[k].quota_2020
    new_uni_set[i].cq_2020p = cq

cq_2017.sort()
cq_2018.sort()
cq_2019.sort()
cq_2020.sort()
ranking_2017.sort()
ranking_2018.sort()
ranking_2019.sort()
ranking_2020.sort()

plt.plot(cq_2017,ranking_2017,color='red')
plt.plot(cq_2018,ranking_2018,color='orange')
plt.plot(cq_2019,ranking_2019,color='blue')
# plt.plot(cq_2020,ranking_2020,color='green')
plt.show()

def CQ_ranking_2017(x):
    if x < cq_2017[1]: return x*ranking_2017[1]/cq_2017[1]
    if x > cq_2017[len(cq_2017)-1]: return ranking_2017[len(cq_2017)-1]
    for i in range(1,len(cq_2017)):
        if x < cq_2017[i]: return ranking_2017[i-1]+(x-cq_2017[i-1])*(ranking_2017[i]-ranking_2017[i-1])/(cq_2017[i]-cq_2017[i-1])

def CQ_ranking_2018(x):
    if x < cq_2018[1]: return x*ranking_2018[1]/cq_2018[1]
    if x > cq_2018[len(cq_2018)-1]: return ranking_2018[len(cq_2018)-1]
    for i in range(1,len(cq_2018)):
        if x < cq_2018[i]: return ranking_2018[i-1]+(x-cq_2018[i-1])*(ranking_2018[i]-ranking_2018[i-1])/(cq_2018[i]-cq_2018[i-1])

def CQ_ranking_2019(x):
    if x < cq_2019[1]: return x*ranking_2019[1]/cq_2019[1]
    if x > cq_2019[len(cq_2019)-1]: return ranking_2019[len(cq_2019)-1]
    for i in range(1,len(cq_2019)):
        if x < cq_2019[i]: return ranking_2019[i-1]+(x-cq_2019[i-1])*(ranking_2019[i]-ranking_2019[i-1])/(cq_2019[i]-cq_2019[i-1])

def CQ_ranking_2020(x):
    if x < cq_2020[1]: return x*ranking_2020[1]/cq_2020[1]
    if x >= cq_2020[len(cq_2020)-1]: return ranking_2020[len(cq_2020)-1]
    for i in range(1,len(cq_2020)):
        if x < cq_2020[i]: return ranking_2020[i-1]+(x-cq_2020[i-1])*(ranking_2020[i]-ranking_2020[i-1])/(cq_2020[i]-cq_2020[i-1])
    print(x)
    return 0

def CQ_ranking_2020p(x):
    return (CQ_ranking_2017(x)+CQ_ranking_2018(x)+CQ_ranking_2019(x))/3

def ranking_score_2017(x):
    for i in range(2,c_table2017.nrows):
        a = int(c_table2017.cell(i,0).value)
        b = int(c_table2017.cell(i,2).value)
        c = int(c_table2017.cell(i-1,0).value)
        d = int(c_table2017.cell(i-1,2).value)
        if x < b:
            return c + (x-d)*(a-c)/(b-d)

def ranking_score_2020(x):
    for i in range(2,c_table2020.nrows):
        a = int(c_table2020.cell(i,0).value)
        b = int(c_table2020.cell(i,2).value)
        c = int(c_table2020.cell(i-1,0).value)
        d = int(c_table2020.cell(i-1,2).value)
        if x < b:
            return c + (x-d)*(a-c)/(b-d)


workbook = xlwt.Workbook()
worksheet2017 = workbook.add_sheet('2017')
worksheet2018 = workbook.add_sheet('2018')
worksheet2019 = workbook.add_sheet('2019')
worksheet2020 = workbook.add_sheet('2020')
tier1 = workbook.add_sheet("tier1")
tier2 = workbook.add_sheet("tier2")

for i in range(0,len(new_uni_set)):
    n = new_uni_set[i].uni_name
    if n[len(n)-6:len(n)] == "平行录取一段":
        cq = new_uni_set[i].cq_2020p
        ranking = CQ_ranking_2020(cq)#
        score = ranking_score_2020(ranking)
        tier1.write(i,0,new_uni_set[i].uni_name)
        tier1.write(i,1,new_uni_set[i].score_line_2020)
        tier1.write(i,2,score)

for i in range(0,len(new_uni_set)):
    n = new_uni_set[i].uni_name
    if n[len(n)-6:len(n)] == "平行录取二段":
        cq = new_uni_set[i].cq_2020p
        ranking = CQ_ranking_2020(cq)#
        score = ranking_score_2020(ranking)
        tier2.write(i,0,new_uni_set[i].uni_name)
        tier2.write(i,1,new_uni_set[i].score_line_2020)
        tier2.write(i,2,score)


for i in range(0,len(new_uni_set)):
    worksheet2017.write(i,0,new_uni_set[i].uni_name)
    worksheet2017.write(i,1,new_uni_set[i].uni_ranking_2017)
    worksheet2017.write(i,2,new_uni_set[i].quota_2017)
for i in range(0,len(new_uni_set)):
    worksheet2018.write(i,0,new_uni_set[i].uni_name)
    worksheet2018.write(i,1,new_uni_set[i].uni_ranking_2018)
    worksheet2018.write(i,2,new_uni_set[i].quota_2018)
for i in range(0,len(new_uni_set)):
    worksheet2019.write(i,0,new_uni_set[i].uni_name)
    worksheet2019.write(i,1,new_uni_set[i].uni_ranking_2019)
    worksheet2019.write(i,2,new_uni_set[i].quota_2019)
for i in range(0,len(new_uni_set)):
    worksheet2020.write(i,0,new_uni_set[i].uni_name)
    worksheet2020.write(i,1,new_uni_set[i].uni_ranking_2020)
    worksheet2020.write(i,2,new_uni_set[i].quota_2020)

workbook.save('浙江ranking.xls')

print("over")