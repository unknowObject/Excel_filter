# coding=utf-8
# author Yunong Zhang

import xlrd
import xlwt
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Combobox
from tkinter.filedialog import askdirectory
import os


# 判断字符串str1是否包含在str2中
def str_in_str(str1, str2):
    if str1 in str2:
        return True
    else:
        return False


def selectfPath():
    path_ = filedialog.askopenfilename()  # 使用askdirectory()方法返回文件夹的路径
    if path_ == "":
        f_path.get()  # 当打开文件路径选择框后点击"取消" 输入框会清空路径，所以使用get()方法再获取一次路径
    else:
        path_ = path_.replace("/", "\\")  # 实际在代码中执行的路径为“\“ 所以替换一下
        f_path.set(path_)


def selecttPath():
    path_ = askdirectory()  # 使用askdirectory()方法返回文件夹的路径
    if path_ == "":
        t_path.get()  # 当打开文件路径选择框后点击"取消" 输入框会清空路径，所以使用get()方法再获取一次路径
    else:
        path_ = path_.replace("/", "\\")  # 实际在代码中执行的路径为“\“ 所以替换一下
        t_path.set(path_)


def opentPath():
    dir = os.path.dirname(t_path.get() + "\\")
    os.system('start ' + dir)
    # print(dir)


# 判断关键字是否在输入inStr中
def str_in_dic(inStr, compare_dic):
    keys = compare_dic.keys()
    for i in keys:
        if i in inStr:
            return compare_dic[i]
    return '未知类别'


# 判断关键字是否在输入inStr中
def str_in_And(inStr, compare_TupDic):
    keys = compare_TupDic.keys()
    for i in keys:
        if i[0] in inStr and i[1] in inStr:
            return compare_TupDic[i]
    return '未知类别'


# 判断关键字是否在输入inStr中
def str_in_Or(inStr, compare_TupDic):
    keys = compare_TupDic.keys()
    for i in keys:
        if i[0] in inStr or i[1] in inStr:
            return compare_TupDic[i]
    return '未知类别'


# 返回分类结果列表，输入config_list为判断条件关键字
def get_result(config_dic, table_sheet, sortM):
    result = ["---占位---"]
    # 以字典形式用所有关键字比对投诉内容，并将所属类别打印，如不存在写不明类别
    index = content_index.get() - 1
    if index <= -1:
        print("Error: target index out of range.")

    if sortM == "包含":
        for i in range(1, table_sheet.nrows):
            try:
                result.append(str_in_dic(table_sheet.cell_value(i, index), config_dic))
            except Exception as e:
                print("Error: " + str(e))
    elif sortM == "包含（与）":
        for i in range(1, table_sheet.nrows):
            try:
                result.append(str_in_And(table_sheet.cell_value(i, index), config_dic))
            except Exception as e:
                print("Error: " + str(e))
    elif sortM == "包含（或）":
        for i in range(1, table_sheet.nrows):
            try:
                result.append(str_in_Or(table_sheet.cell_value(i, index), config_dic))
            except Exception as e:
                print("Error: " + str(e))

    print("\n分类内容所在列数：" + str(content_index.get()))
    print("\n分类结果：")
    for i in range(len(result)):
        print(str(i) + ": " + result[i])
    return result


# 整理输入config文件为字典格式
def get_config_dic():
    # 打开配置文件
    config_file = open('config/config_contain.txt', 'r', encoding='utf-8')
    config_content = config_file.readlines()
    # print(config_content)
    print("\n配置文件：")

    # 将关键词写入config_dic中，‘Name’: 'Category'
    config_dic = {}
    for i in config_content:
        index1 = i.find('\t')
        index2 = i[index1].find('\n')
        name = i[:index1]
        category = i[index1 + 1:index2]
        config_dic[name] = category
        print(name + ":\t" + category)

    # print(config_dic)
    return config_dic


# 整理输入config文件为字典格式
def get_config_dicAnd():
    # 打开配置文件
    config_file = open('config/config_and.txt', 'r', encoding='utf-8')
    config_content = config_file.readlines()
    # print(config_content)
    print("\n配置文件：")

    # 将关键词写入config_dic中，‘Name’: 'Category'
    config_dic = {}
    for i in config_content:
        index1 = i.find('\t')
        index2 = i[index1 + 1:].find('\t')
        index3 = i[index2].find('\n')
        name1 = i[:index1]
        name2 = i[index1 + 1:index1 + index2 + 1]
        category = i[index1 + index2 + 2:index3]
        nameTup = (name1, name2)
        config_dic[nameTup] = category
        print(nameTup[0] + " " + nameTup[1] + ":\t" + category)

    # print(config_dic)
    return config_dic


# 整理输入config文件为字典格式
def get_config_dicOr():
    # 打开配置文件
    config_file = open('config/config_or.txt', 'r', encoding='utf-8')
    config_content = config_file.readlines()
    # print(config_content)
    print("\n配置文件：")

    # 将关键词写入config_dic中，‘Name’: 'Category'
    config_dic = {}
    for i in config_content:
        index1 = i.find('\t')
        index2 = i[index1 + 1:].find('\t')
        index3 = i[index2].find('\n')
        name1 = i[:index1]
        name2 = i[index1 + 1:index1 + index2 + 1]
        category = i[index1 + index2 + 2:index3]
        nameTup = (name1, name2)
        config_dic[nameTup] = category
        print(nameTup[0] + " " + nameTup[1] + ":\t" + category)

    # print(config_dic)
    return config_dic


# 读取输入文件，配置文件，分类后写入目标文件
def open_write():
    sort_method = var.get()
    print(sort_method)
    # 获取输入文件路径
    in_file_path = f_path.get()
    print("Input file : " + in_file_path)
    # 打开目标excel表格
    try:
        table_file = xlrd.open_workbook(in_file_path)
    except Exception as e:
        print(e)

    # 找到对应列数
    table_sheet = table_file.sheet_by_name(sheet_name)

    # 获取配置文件信息
    if sort_method == "包含":
        config_dic = get_config_dic()
    elif sort_method == "包含（与）":
        config_dic = get_config_dicAnd()
    elif sort_method == "包含（或）":
        config_dic = get_config_dicOr()

    result = get_result(config_dic, table_sheet, sort_method)

    # 获取输出文件路径
    out_file_path = t_path.get()
    print("\nOutput file: " + out_file_path + "\\result.xls")
    # 将结果输出至目标文件
    result_workbook = xlwt.Workbook()
    result_sheet = result_workbook.add_sheet('result_sheet')
    for i in range(len(result)):
        result_sheet.write(i, 0, result[i])
    result_workbook.save(out_file_path + "/result.xls")
    return


if __name__ == '__main__':
    # 投诉内容所在列数
    sheet_name = 'Sheet1'

    # 设置窗口参数
    root = tk.Tk()
    root.geometry('640x480')
    root.title('Tinker')
    root.resizable(False, False)

    # 设置画布背景色
    canvas_root = tk.Canvas(root, bg='white', width=640, height=480)
    canvas_root.pack()

    # 设置logo大小以及位置
    im_root1 = PhotoImage(file='images/anvil&hammer_insoftware.GIF')
    img_label = Label(root, image=im_root1)
    img_label.place(x=215, y=20, width=195, height=180)

    # 选择文件
    f_path = StringVar()
    f_path.set(os.path.abspath("."))

    Label(root, bg='white', text="文件路径:").place(x=20, y=220, width=60, height=30)
    Entry(root, textvariable=f_path).place(x=80, y=220, width=400, height=30)

    # e.insert(0,os.path.abspath("."))
    Button(root, text="文件选择(.xls)", command=selectfPath).place(x=480, y=220, width=140, height=30)

    # 选择目标目录
    t_path = StringVar()
    t_path.set(os.path.abspath("."))

    Label(root, bg='white', text="目标路径:").place(x=20, y=280, width=60, height=30)
    Entry(root, textvariable=t_path).place(x=80, y=280, width=400, height=30)

    # e.insert(0,os.path.abspath("."))
    Button(root, text="路径选择", command=selecttPath).place(x=480, y=280, width=65, height=30)
    Button(root, text="打开文件夹", command=opentPath).place(x=545, y=280, width=75, height=30)

    content_index = IntVar()
    Label(root, bg='white', text="分类内容所在列数:").place(x=20, y=320, width=100, height=30)
    Entry(root, textvariable=content_index).place(x=130, y=320, width=30, height=30)

    var = tk.StringVar()
    Combobox(root, state='readonly', textvariable=var, values=('包含', '包含（与）', '包含（或）')).place(x=20, y=350,
                                                                                                    width=100,
                                                                                                    height=30)
    Button(root, text="开始分类", command=open_write).place(x=282, y=350, width=80, height=50)
    Label(root, bg='white', text="v1.0.3").place(x=580, y=450, width=60, height=30)

    mainloop()
