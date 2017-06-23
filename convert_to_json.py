#!/usr/bin/python3
# -*- coding: UTF-8 -*-

import xlrd,json

wb = xlrd.open_workbook("一级.xls")
sheet0 = wb.sheet_by_index(0)

# 获得行数和列数
nrows = sheet0.nrows    # 行总数
ncols = sheet0.ncols    # 列总数

fo = open("一级.txt", "w+")

questions = []

# 判断选择还是填空，并分别执行
for i in range(1, nrows):
    topic = sheet0.cell_value(i, 3)
    answer = sheet0.cell_value(i, 9)
    answer = answer.replace('A','1').replace('B','2').replace('C','3').replace('D','4')

    q_type = sheet0.cell_value(i,2)
    if q_type == "选择":
        option1 = str(sheet0.cell_value(i, 10))
        option2 = str(sheet0.cell_value(i, 11))
        option3 = str(sheet0.cell_value(i, 12))
        option4 = str(sheet0.cell_value(i, 13))

        question = str(i) + '.' + topic
        option = [option1,option2,option3,option4]
        q_dict = {'question':question,'answer':option,'correctAnswer':int(answer)}
        questions.append(q_dict)

questions = {'questions':questions}
print(json.dumps(questions,ensure_ascii=False))


fo.close()


