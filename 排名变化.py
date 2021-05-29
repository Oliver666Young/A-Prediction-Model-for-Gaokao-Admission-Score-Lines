import xlrd
file = '新程序\大学排名变化.xlsx'

wb = xlrd.open_workbook(filename=file)
rank_2017 = wb.sheet_by_index(0)
rank_2020 = wb.sheet_by_index(7)

b=0
for k in range(1,1719):
    a=0
    for i in range(1,541):
        if rank_2017.cell_value(k,0)==rank_2020.cell_value(i,0):
            print(rank_2020.cell_value(i,1))
            a=1
            b=b+1
            break
    if a == 0:
        print (0)
print(b)