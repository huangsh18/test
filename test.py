import xlrd
import xlwt
data1 = xlrd.open_workbook('popandord.xlsx')
table1 = data1.sheets()[0]
nr = table1.nrows ##229
rows=[]
dict1={}
dict2={}
dict3={}
for i in range(1,229):
    rows.append(table1.row_values(i))   #每一行数据加入rows
    #print(rows[i-1])
    dict1[rows[i-1][0]] = [rows[i-1][1],rows[i-1][2]]
    #print(dict1)   #城市的人口和快递量


data2 = xlrd.open_workbook('distance.xlsx')
table2 = data2.sheets()[0]
for i in range(1,229):
    dict2[table2.col_values(0)[i]] = i    #距离表里面城市对应编号
    dict3[i] = table2.col_values(0)[i]    #距离表里面编号对应城市
#print(dict2)
#print(dict3)


def dc(i,j):
    return float(dict1[dict3[i]][0]*dict1[dict3[i]][1]*dict1[dict3[j]][0]*dict1[dict3[j]][1])   #分子下面的计算


workbook = xlwt.Workbook(encoding = 'utf-8')
worksheet = workbook.add_sheet('count')
for i in range(1,229):
    for j in range(1,229):
        worksheet.write(i, j, label=pow(table2.cell(i,j).value,2)/pow(dc(i,j),0.05))   #把计算后的结果写入表格得到d’（xi，xj）
workbook.save('counttest.xls')

cols=[]
data3 = xlrd.open_workbook('counttest.xls')
table3 = data3.sheets()[0]
for i in range(1,229):
    cols.append(table3.col_values(i))  #读取每一列的数据  cols城市编号从0开始
cols_new=[]
mid_value = []
for i in range(0,228):
    cols[i].remove('')
    cols_new.append(sorted(cols[i]))
    lens = len(cols[i])
    if lens%2 == 0:   #取得列向量中间值
        mid_value.append((cols_new[i][int(lens/2)]+cols_new[i][int(lens/2-1)])/2)
    else:
        mid_value.append(cols_new[i][int(lens/2)])
print(mid_value)
e = sum(mid_value)/len(mid_value)
print(e)
mp=[]
for i in range(0,228):
    count = 0
    for j in cols[i]:
        if j <= e:
            count = count + 1
    mp.append(count)
print(mp)  #每个城市有多少个符合e距离的城市  ，mp的城市编号从0开始
MinPts = sum(mp)/(2*len(mp))
print("MinPts",MinPts)
G = []
for i in range(0,228):
    if mp[i] >= MinPts:
        G.append(i)
print("G",G)  #符合MinPts距离个数内城市的编号 0开始的，G表示符合的城市的数组

workbook = xlwt.Workbook(encoding = 'utf-8')
worksheet = workbook.add_sheet('G_city')
for i in range(1,229):
    worksheet.write(0, i, label=dict3[i])
    worksheet.write(i, 0, label=dict3[i])

for i in range(0,228):
    for j in range(0,228):
        if i in G and j in G:
            worksheet.write(i+1,j+1,label=table3.cell(i+1,j+1).value)
        else:
            worksheet.write(i+1,j+1,label=0)
workbook.save('G_city.xls')   #把G的计算的距离表写出来

c=[]
c.append(mp.index(max(mp)))  #符合e的


# 最多的城市的编号
print(c)  #c代表的城市号是从0起

data4 = xlrd.open_workbook('G_city.xls')
table4 = data4.sheets()[0]

while len(c)<len(G):
    summax = []
    for j in range(1,229):  #列
        sum1 = 0
        for i in range(0, len(c)):  #属于C的行
            sum1 = sum1 + table4.cell(c[i]+1, j).value
        summax.append(sum1)
    c.append(summax.index(max(summax)))
print('C',c)
new_c =[]
for ic in c:
    if ic not in new_c:
        new_c.append(ic)
print(new_c)
dict4={}
for i in new_c:
    dict4[i+1]=[]
print(dict4)

for i in range(1,229):
    fnum = []
    for j in new_c:
        fnum.append(table3.cell(i,j+1).value)
    indexnum = fnum.index(min(fnum))
    dict4[new_c[indexnum]+1].append(i)
print(dict4)
