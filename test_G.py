import xlwt
import xlrd

dict2={}
dict3={}
data2 = xlrd.open_workbook('distance.xlsx')
table2 = data2.sheets()[0]
for i in range(1,229):
    dict2[table2.col_values(0)[i]] = i    #距离表里面城市对应编号
    dict3[i] = table2.col_values(0)[i]    #距离表里面编号对应城市

data3 = xlrd.open_workbook('counttest.xls')
table3 = data3.sheets()[0]

workbook = xlwt.Workbook(encoding = 'utf-8')
worksheet = workbook.add_sheet('g_city_test')
for i in range(1,229):
    worksheet.write(0, i, label=dict3[i])
    worksheet.write(i, 0, label=dict3[i])

G=[1,3,5,7,8,9,11,13]

for i in G:
    for j in G:
        worksheet.write(i+1,j+1,label=table3.cell(i+1,j+1).value)
workbook.save('G_city_test.xls')   #把G的计算的距离表写出来