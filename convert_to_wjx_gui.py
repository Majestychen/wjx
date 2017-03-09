#!/usr/bin/python3
# -*- coding: UTF-8 -*-

import xlrd
from tkinter import *
from tkinter import filedialog

def convert(inputfile,outputfile):
    wb = xlrd.open_workbook(inputfile)
    sheet0 = wb.sheet_by_index(0)

    # 获得行数和列数
    nrows = sheet0.nrows    # 行总数
    ncols = sheet0.ncols    # 列总数

    #分解题目函数
    def braek_topic(topic,answer,i):
        #尝试拆题，失败报错
        try:
            Lbox.insert(END, "第%s题需要拆解，正在拆解..." % i)
            topic = topic.replace("(", "（")
            topic = topic.replace(")","）")
            topic = topic.replace("（）", "（ ）")
            topic = topic.replace("\u3000",'')  #解决编码问题
            topic = topic.split("（ ",1)
            print(len(topic),topic)
            #Lbox.insert(END, len(topic))
            #Lbox.insert(END,topic)
            topic = topic[0] + '（ ' + answer + topic[1]
            Lbox.insert(END,"...成功拆解")
            return topic
        except:
            Lbox.insert(END,"拆解第%s题失败，请检查题库文件"%i)
            Lbox.see(END)
            return None


    fo = open(outputfile, "w")

    #判断选择／多选／填空，并分别执行
    for i in range(1, nrows):
        Lbox.insert(END,">>>正在解析第%s题"%i)
        topic = sheet0.cell_value(i, 3)
        answer = sheet0.cell_value(i, 9)

        #判断题
        if sheet0.cell_value(i, 12)=='':
            option1 = str(sheet0.cell_value(i, 10))
            option2 = str(sheet0.cell_value(i, 11))

            if answer == "A":
                # 写入文件
                fo.write(str(i) + '、' + topic  + " (对)"+'\n')
                fo.write("A." + option1 + '\n')
                fo.write("B." + option2 + '\n\n')
            else:
                # 写入文件
                fo.write(str(i) + '、' + topic + " (错)" + '\n')
                fo.write("A." + option1 + '\n')
                fo.write("B." + option2 + '\n\n')

        #选择题
        else:

            #单选
            if len(answer) > 1:

                option1 = str(sheet0.cell_value(i, 10))
                option2 = str(sheet0.cell_value(i, 11))
                option3 = str(sheet0.cell_value(i, 12))
                option4 = str(sheet0.cell_value(i, 13))

                #写入文件
                topic = braek_topic(topic, answer,i)
                fo.write(str(i) + '、' + topic + ' [多选题]' +'\n')
                fo.write("A." + option1+'\n')
                fo.write("B." + option2+'\n')
                fo.write("C." + option3+'\n')
                fo.write("D." + option4+'\n\n')

            #多选
            else :
                option1 = str(sheet0.cell_value(i, 10))
                option2 = str(sheet0.cell_value(i, 11))
                option3 = str(sheet0.cell_value(i, 12))
                option4 = str(sheet0.cell_value(i, 13))

                # 写入文件
                topic = braek_topic(topic, answer,i)
                fo.write(str(i) + '、' + topic + ' [单选题]' + '\n')
                fo.write("A." + option1 + '\n')
                fo.write("B." + option2 + '\n')
                fo.write("C." + option3 + '\n')
                fo.write("D." + option4 + '\n\n')

    fo.close()

    Lbox.insert(END,"已完成"+str(nrows-1)+"道题目排编!")

def testfuc():
    a = E1.get()
    b = "output.txt"
    convert(a,b)
    #print("done!")
    Lbox.see(END)

def getfilepath():
    filepath = filedialog.askopenfilename()
    E1.delete(0,END)
    E1.insert(END,filepath)
    print(filepath)

#GUI
root = Tk()
root.title('题库转换器')
#root.geometry("380x260+500+500")
Label(root,text = "输入题库文件路径：").pack()
E1 = Entry(root,width=40,bd=3)
E1.pack()
Button(root,text = "选择文件",command = getfilepath).pack() #加入文件选择窗
Button(root,text = "转换",fg = 'black',bg = 'green',width=7,command = testfuc).pack()
Lbox = Listbox(root,width=40)
Lbox.pack()
Label(root,text="新媒体工作室",anchor=CENTER).pack()


root.mainloop()
