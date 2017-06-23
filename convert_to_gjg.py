#!/usr/bin/python3
# -*- coding: UTF-8 -*-

import xlrd

wb = xlrd.open_workbook("/Users/guo/project/conver/题库/《抽考题库》钢轨探伤工.xls")
sheet0 = wb.sheet_by_index(0)

# 获得行数和列数
nrows = sheet0.nrows  # 行总数
ncols = sheet0.ncols  # 列总数


# 分解题目函数
def braek_topic(topic, answer):
    topic = topic.replace("\u3000", " ")
    topic = topic.replace("(", "（")
    topic = topic.replace(")", "）")
    topic = topic.replace("（）", "（ ）")
    topic_list = topic.split("（ ", 1)
    # 解决没有括号的问题
    if len(topic_list) == 2:
        topic = topic_list[0] + '（ ' + answer + topic_list[1]
        print(len(topic_list), topic)
        return topic
    else:
        topic = topic + "( )"
        topic = topic.replace("\u3000", " ")
        topic = topic.replace("(", "（")
        topic = topic.replace(")", "）")
        topic = topic.replace("（）", "（ ）")
        topic_list = topic.split("（ ", 1)
        topic = topic_list[0] + '（ ' + answer + topic_list[1]
        print(len(topic_list), topic)
        return topic


fo = open("output.txt", "w")

# 判断选择／多选／填空，并分别执行
for i in range(1, nrows):
    topic = sheet0.cell_value(i, 3)
    answer = sheet0.cell_value(i, 9)

    # 判断题
    if sheet0.cell_value(i, 12) == '':
        option1 = str(sheet0.cell_value(i, 10))
        option2 = str(sheet0.cell_value(i, 11))

        if answer == "A":
            # 写入文件
            fo.write(topic + " (对)" + '\n\n')
        else:
            # 写入文件
            fo.write(topic + " (错)" + '\n\n')

    # 选择题
    else:

        # 单选
        if len(answer) > 1:

            option1 = str(sheet0.cell_value(i, 10))
            option2 = str(sheet0.cell_value(i, 11))
            option3 = str(sheet0.cell_value(i, 12))
            option4 = str(sheet0.cell_value(i, 13))

            # 写入文件
            topic = braek_topic(topic, answer)
            fo.write(topic + ' [多选题]' + '\n')
            fo.write("A." + option1 + '\n')
            fo.write("B." + option2 + '\n')
            fo.write("C." + option3 + '\n')
            fo.write("D." + option4 + '\n\n')

        # 多选
        else:
            option1 = str(sheet0.cell_value(i, 10))
            option2 = str(sheet0.cell_value(i, 11))
            option3 = str(sheet0.cell_value(i, 12))
            option4 = str(sheet0.cell_value(i, 13))

            # 写入文件
            topic = braek_topic(topic, answer)
            fo.write(topic + ' [单选题]' + '\n')
            fo.write("A." + option1 + '\n')
            fo.write("B." + option2 + '\n')
            fo.write("C." + option3 + '\n')
            fo.write("D." + option4 + '\n\n')

fo.close()

print("已完成" + str(nrows - 1) + "道题目排编")
