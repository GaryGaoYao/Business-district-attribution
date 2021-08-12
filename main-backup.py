import PositionDesition
import openpyxl
import numpy as np
import matplotlib.pyplot as plt
import shangquanduqu

workbook1 = openpyxl.load_workbook('G:\\python+codes\position+zhangxiyu\\xiyu\\tr-0812.xlsx')
worksheet1 = workbook1.active
minrow1 = worksheet1.min_row  # 最小行
maxrow1 = worksheet1.max_row  # 最大行

pointlist_market = []
PointListMap = []
shangquan_temporary_list = []
x = []
y = []

# 全量商户的经纬度度读取
for n in range(minrow1, maxrow1 + 1):
    cell_jingdu = worksheet1.cell(n, 18).value
    cell_weidu = worksheet1.cell(n, 19).value
    pointlist_market.append((cell_jingdu, cell_weidu))

for n in range(1, 164):
    shangquan_name_value = shangquanduqu.namelist[n-1]
    shangquan_temporary_list_start = shangquanduqu.partitionlist[n-1]
    shangquan_temporary_list_end = shangquanduqu.partitionlist[n]

    for i in range(shangquanduqu.partitionlist[n-1],shangquanduqu.partitionlist[n]):
        shangquan_temporary_list.append(shangquanduqu.pointlist_market[i])
        if i == shangquanduqu.partitionlist[n]-1:
            #判断全量商户坐标
            Insert_column_number = 60
            for j in range(0, len(pointlist_market)):
                a_x, a_y = pointlist_market[j]
                if (PositionDesition.IsPtInPoly(a_x, a_y, shangquan_temporary_list)):
                    worksheet1.cell(j + 1, Insert_column_number).value = shangquan_name_value

    #清空零时列表
    shangquan_temporary_list = []

#最后一个补充漏下的商圈
guizhoudasha_tongzhou = [(111，111),(111，111),(111，111),(111，111)]
for j in range(0, len(pointlist_market)):
    a_x, a_y = pointlist_market[j]
    if (PositionDesition.IsPtInPoly(a_x, a_y, guizhoudasha_tongzhou)):
        worksheet1.cell(j + 1, 60).value = 'XXXX'

workbook1.save('tr-0812_result.xlsx')
