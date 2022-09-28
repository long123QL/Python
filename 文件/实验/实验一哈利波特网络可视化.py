import openpyxl
from matplotlib import pyplot as plt
import networkx as nx


def matching(id, dict_all):  # 字典匹配文字
    for dict in dict_all.keys():
        if (id == dict):
            return dict_all[dict]


def deltel(listX):  # 删除重复的字段
    res = []
    for i in listX:
        if i not in res:
            res.append(i)
    return res


def draw(nodes, edges, num):  # 调用画图
    str = "D:\学_Pycharm文件\文件\文件\网络试图{0}.png".format(num)  # 生成文件路径
    G = nx.Graph()
    G.add_nodes_from(nodes)
    G.add_edges_from(edges)
    plt.figure(figsize=(40, 30))
    nx.draw_networkx(G)
    plt.savefig(str, format="PNG")
    plt.show()


filename = openpyxl.load_workbook('D:\学_Pycharm文件\文件\文件\harry_potter .xlsx')  # 打开文件
sheet1 = filename['relation']  # 获取表relation
sheet2 = filename['character']  # 获取表character
rowR = sheet1.rows  # 读取relation表
rowS = sheet2.rows  # 读取character的单元格数据
cases1 = []
cases2 = []
for r in list(rowS)[1:]:  # 遍历表格所有行
    case = []
    for i in r[:1]:  # 遍历第一个的单元格，存入case1
        cases1.append(i.value)
    for i in r[1:2]:  # 遍历第二个的单元格，存入case2
        cases2.append(i.value)
dict_all = dict(zip(cases1, cases2))  # 生成字典
print(dict_all)

tup_three = []
tup_all = []
for r in list(rowR)[1:]:  # 遍历表格所有行
    Three_pepple = []
    All_pepple = []
    for i in r:  # 遍历每一行的单元格，存入case
        valuez = matching(i.value, dict_all)  # 调用文字匹配
        All_pepple.append(valuez)  # 生成所有人的文字关系
        Three_pepple.append(i.value)  # 生成主角关系列表
    tup_three.append(tuple(Three_pepple))
    tup_all.append(tuple(All_pepple))
print(tup_three)
print(tup_all)

column = sheet1.columns  # 读取所有人ID
All_pepple_id = []
for i in list(column)[0]:  # 读取列信息
    valuez = matching(i.value, dict_all)  # 文字ID
    All_pepple_id.append(valuez)
del All_pepple_id[0]  # 删除表头
id_all = deltel(All_pepple_id)  # 调用删除重复
print(id_all)

list_three = []
for items in tup_three[0:]:  # 生成主角数字关系
    for itemsnodes in items:
        if (itemsnodes == 21 or itemsnodes == 39 or itemsnodes == 58):
            list_three.append(items)
print(list_three)
three_gonist_id = []
for i in range(len(list_three)):  # 处理元组成列表
    p = lambda x: x[i][0]
    three_gonist_id.append(p(list_three))
print(three_gonist_id)
id_three = deltel(three_gonist_id)
print(id_three)

List_three = []  # 生成id文字关系
for i in id_three:
    valuez = matching(i, dict_all)  # 调用匹配
    List_three.append(valuez)
print(List_three)

tup_three = []
for item in list_three[1:]:  # 生成文字关系
    case = []
    for itemsnode in item:
        valuez = matching(itemsnode, dict_all)  # 调用匹配
        case.append(valuez)
    tup_three.append(tuple(case))
print(tup_three)

draw(id_all, tup_all, 1)  # 调用ALL画笔
draw(List_three, tup_three, 2)  # 调用Three画笔
