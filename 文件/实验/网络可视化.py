import openpyxl
from matplotlib import pyplot as plt
import networkx as nx
filename = openpyxl.load_workbook('D:\学_Pycharm文件\文件\文件\harry_potter .xlsx')  #打开文件
sheet1 =filename ['relation']       #获取表relation
sheet2 =filename ['character']       #获取表character
rows = sheet2.rows       # 读取character的单元格数据

cases1=[]
cases2=[]
for r in list(rows)[1:]:  # 遍历表格所有行
    case = []
    for i in r[:1]:  # 遍历第一个的单元格，存入case1
        cases1.append(i.value)
    for i in r[1:2]:  # 遍历第二个的单元格，存入case2
        cases2.append(i.value)
dict_all = dict(zip(cases1, cases2))   #生成字典
print(dict_all)
rows = sheet1.rows      #读取relation表
cases =[]
tup=[]
tup1=[]
for r in list(rows)[1:]:  # 遍历表格所有行
    case = []
    for i in r:  # 遍历每一行的单元格，存入case
        # for dict in dict_all.keys():  # 匹配文字
        #     if (i.value == dict):
        #         case.append(dict_all[dict])
        case.append(i.value)
    # tup.append(tuple(case))
    tup1.append(tuple(case))
# print(tup1)

listX=[]
for items in tup1[0:]:
    for itemsnodes in items:
        if(itemsnodes==21 or itemsnodes==39 or itemsnodes==58):
            listX.append(items)

b=[]
for i in range(len(listX)):
    p=lambda x:x[i][0]
    b.append(p(listX))
print(b)
res = []
for i in b:   #删除重复的字段
    if i not in res:
        res.append(i)
print(res)

d=[]
for i in res:
    for dict in dict_all.keys():
        if (i == dict):
            d.append(dict_all[dict])

for item in listX[1:]:
    case=[]
    for itemsnode in item:
        for dict in dict_all.keys():  # 匹配文字
            if (itemsnode == dict):
                case.append(dict_all[dict])
    tup.append(tuple(case))
print(tup)


# column = sheet1.columns
# titles=[]
# for i in list(column)[0]:      #读取列信息
#     for dict in dict_all.keys():
#         if (i.value == dict):
#             titles.append(dict_all[dict])
# del titles[0]   #删除表头
# res = []
# for i in titles:   #删除重复的字段
#     if i not in res:
#         res.append(i)
# print(res)
#
G=nx.Graph()
G.add_nodes_from(d)
G.add_edges_from(tup)
plt.figure(figsize = (40,30))
nx.draw_networkx(G)
plt.show()


