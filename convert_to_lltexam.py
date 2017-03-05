#!/usr/bin/python3
# -*- coding: UTF-8 -*-

import xlrd

wb = xlrd.open_workbook("一级.xls")
sheet0 = wb.sheet_by_index(0)

# 获得行数和列数
nrows = sheet0.nrows    # 行总数
ncols = sheet0.ncols    # 列总数

fo = open("一级.txt", "w+")

# 判断选择还是填空，并分别执行
for i in range(1, nrows):
    topic = sheet0.cell_value(i, 3)
    answer = sheet0.cell_value(i, 9)

    #判断
    if sheet0.cell_value(i, 12)=='':
        option1 = str(sheet0.cell_value(i, 10))
        option2 = str(sheet0.cell_value(i, 11))

        print(str(i) + '.' + topic)
        print(answer)
        print("A." + option1)
        print("B." + option2)

        # 写入文件
        fo.write("【判断题】"+str(i) + '.' + topic+'\n')
        fo.write(answer+'\n')
        fo.write("A." + option1+'\n')
        fo.write("B." + option2+'\n')

    else:
        # 选择
        if len(answer) > 1:
            # 多选
            option1 = str(sheet0.cell_value(i, 10))
            option2 = str(sheet0.cell_value(i, 11))
            option3 = str(sheet0.cell_value(i, 12))
            option4 = str(sheet0.cell_value(i, 13))

            print(str(i) + '.' + topic)
            print(answer)
            print("A." + option1)
            print("B." + option2)
            print("C." + option3)
            print("D." + option4)

            # 写入文件
            fo.write("【多选题】"+str(i) + '.' + topic+'\n')
            fo.write(answer+'\n')
            fo.write("A." + option1+'\n')
            fo.write("B." + option2+'\n')
            fo.write("C." + option3+'\n')
            fo.write("D." + option4+'\n')
        else:
            # 单选
            option1 = str(sheet0.cell_value(i, 10))
            option2 = str(sheet0.cell_value(i, 11))
            option3 = str(sheet0.cell_value(i, 12))
            option4 = str(sheet0.cell_value(i, 13))

            print(str(i) + '.' + topic)
            print(answer)
            print("A." + option1)
            print("B." + option2)
            print("C." + option3)
            print("D." + option4)

            # 写入文件
            fo.write("【单选题】" + str(i) + '.' + topic + '\n')
            fo.write(answer + '\n')
            fo.write("A." + option1 + '\n')
            fo.write("B." + option2 + '\n')
            fo.write("C." + option3 + '\n')
            fo.write("D." + option4 + '\n')


fo.close()

print("已完成"+str(nrows-1)+"道题目排编")
