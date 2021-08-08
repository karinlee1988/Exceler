#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2021/1/1 18:45
# @Author : 李加林
# @FileName : VLOOKUP工具.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976
# @github : https://github.com/karinlee1988/
# @gitee : https://gitee.com/karinlee/
# @Personal website : https://karinlee.cn/

import os
import openpyxl
import tkinter as tk
import tkinter.filedialog
from exceler import vlookup

class VlookupGui(object):
    """
    python实现vlookup图形工具

    20210808 test OK
    """
    def __init__(self):
        """
        创建界面
        """
        # 新建窗口
        self.master = tk.Tk()
        # 在界面顶部添加横幅图片
        self.photo = tk.PhotoImage(file="images\\doge2.gif")
        # self.path 用于存放选择的文件路径
        # self.flag 当程序运行完成后给用户提示信息
        # 注意！这些都是tk.StringVar()对象，不是str，其他地方要用的话要用get()方法获取str
        self.temp_path = tk.StringVar()
        self.source_path = tk.StringVar()
        self.flag = tk.StringVar()
        self.v1 = tk.StringVar()  # 主工作薄工作表序号
        self.v2 = tk.StringVar()   # 主工作薄用于查找的列
        self.v3 = tk.StringVar()   # 主工作薄要填写的列
        self.v4 = tk.StringVar()   # 数据工作薄工作表序号
        self.v5 = tk.StringVar()  # 数据工作薄对应主工作薄的列
        self.v6 = tk.StringVar()  # 数据工作薄提供数据列
        self.v7 = tk.StringVar()  # 从第几行开始vlookup
        # 窗口图片横幅
        self.img_pack()
        # 主界面布局
        self.window()
        # 框架布局
        self.frame()
        # 保持运行
        self.master.mainloop()

    def img_pack(self):
        """
        窗口图片横幅
        """
        # 图片贴上去
        img_lable = tk.Label(self.master, image=self.photo)
        img_lable.pack()

    def window(self):
        """
        主界面布局设置
        """
        # 调整窗口默认大小及在屏幕上的位置
        self.master.geometry("800x820+550+20")
        # 窗口的标题栏，自己修改
        self.master.title("vlookup工具by李加林v1.0")
        # 把标题贴上去，自己修改
        tk.Label(self.master,text="VLOOKUP工具",font=("黑体",18)).pack()
        tk.Label(self.master, text="-----------------------", font=("黑体", 16)).pack()

    def frame(self):
        """
        框架布局设置

        """
        # 框架贴上去，再在框架里添加Lable，Entry，Button等控件
        frame1 = tk.Frame(self.master)
        frame1.pack()
        frame2 = tk.Frame(self.master)
        frame2.pack()
        # 输入框，标记，按键
        tk.Label(frame1, text="目标路径:", font=("黑体", 14)).grid(row=1, column=0)
        tk.Label(frame1, text="主工作薄->", font=("黑体", 14)).grid(row=2, column=0)
        tk.Entry(frame1, textvariable=self.temp_path, width=50).grid(row=2, column=1)
        tk.Label(frame1, text="数据工作薄->", font=("黑体", 14)).grid(row=3, column=0)
        tk.Entry(frame1, textvariable=self.source_path, width=50).grid(row=3, column=1)
        tk.Button(frame1, text="选择主工作薄", command=self.select_temp_path, font=("黑体", 14)).grid(row=2, column=2)
        tk.Button(frame1, text="选择数据工作薄", command=self.select_source_path, font=("黑体", 14)).grid(row=3, column=2)

        tk.Label(frame2, text="主工作薄工作表序号（从1开始）:", font=("黑体", 14)).grid(row=1, column=0)
        tk.Entry(frame2, textvariable=self.v1).grid(row=1, column=1)
        tk.Label(frame2, text="主工作薄用于查找的列名(大写字母):", font=("黑体", 14)).grid(row=2, column=0)
        tk.Entry(frame2, textvariable=self.v2).grid(row=2, column=1)
        tk.Label(frame2, text="主工作薄要填写的列名(大写字母):", font=("黑体", 14)).grid(row=3, column=0)
        tk.Entry(frame2, textvariable=self.v3).grid(row=3 ,column=1)
        tk.Label(frame2, text="数据工作薄工作表序号（从1开始）:", font=("黑体", 14)).grid(row=4, column=0)
        tk.Entry(frame2, textvariable=self.v4).grid(row=4, column=1)
        tk.Label(frame2, text="数据工作薄对应主工作薄的列名(大写字母):", font=("黑体", 14)).grid(row=5, column=0)
        tk.Entry(frame2, textvariable=self.v5).grid(row=5, column=1)
        tk.Label(frame2, text="数据工作薄提供数据列名(大写字母):", font=("黑体", 14)).grid(row=6, column=0)
        tk.Entry(frame2, textvariable=self.v6).grid(row=6 ,column=1)
        tk.Label(frame2, text="从第几行开始:", font=("黑体", 14)).grid(row=7, column=0)
        tk.Entry(frame2, textvariable=self.v7).grid(row=7, column=1)

        # 按这个按钮执行主程序
        tk.Button(frame2, text="开始VLOOKUP", command=self.main, font=("黑体", 14)).grid(row=8, column=1,pady=33)
        tk.Entry(frame2, textvariable=self.flag,state="readonly").grid(row=9, column=1)

    def select_temp_path(self):
        """
        选择文件，并获取文件的绝对路径

        """
        # 选择文件，path_select变量接收文件地址
        # 注意：self.path 是tk.StringVar()对象，而path_select是str变量
        path_select = tkinter.filedialog.askopenfilename()
        # 通过replace函数替换绝对文件地址中的/来使文件可被程序读取
        # 注意：\\转义后为\，所以\\\\转义后为\\
        path_select = path_select.replace("/", "\\\\")
        # self.path设置path_select的值
        self.temp_path.set(path_select)
        self.flag.set("准备进行vlookup...")

    def select_source_path(self):
        """
        选择文件，并获取文件的绝对路径

        """
        # 选择文件，path_select变量接收文件地址
        # 注意：self.path 是tk.StringVar()对象，而path_select是str变量
        path_select = tkinter.filedialog.askopenfilename()
        # 通过replace函数替换绝对文件地址中的/来使文件可被程序读取
        # 注意：\\转义后为\，所以\\\\转义后为\\
        path_select = path_select.replace("/", "\\\\")
        # self.path设置path_select的值
        self.source_path.set(path_select)

    def main(self):
        """
        程序入口
        """
        wb_template = openpyxl.load_workbook(self.temp_path.get())
        wb_source = openpyxl.load_workbook(self.source_path.get())

        ws_template_index = int(self.v1.get()) - 1
        template_key = self.v2.get()
        template_value = self.v3.get()

        ws_source_index = int(self.v4.get()) - 1
        source_key = self.v5.get()
        source_value = self.v6.get()
        line = self.v7.get()
        # 进行vlookup处理
        vlookup(wb_template,ws_template_index,template_key,template_value,wb_source,ws_source_index,source_key,source_value,line)
        wb_template.save(os.path.basename(self.temp_path.get()).replace('.xlsx','')+'_已进行vlookup.xlsx')
        # 标志设置为处理完成
        self.flag.set("处理完成！")

if __name__ == '__main__':
    # 运行
    app = VlookupGui()