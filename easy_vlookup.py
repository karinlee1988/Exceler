#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2022/7/12 20:04
# @Author : karinlee
# @FileName : easy_vlookup.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976
# @github : https://github.com/karinlee1988/
# @gitee : https://gitee.com/karinlee/
# @Personal website : https://karinlee.cn/


import openpyxl

from openpyxl.utils import column_index_from_string  # ,get_column_letter

def easyvlookup(
        wb_main,
        ws_main_index:int,
        main_key:str,
        main_value:str,
        wb_source,
        ws_source_index:int,
        source_key:str,
        source_value:str,
        line:int
        ) -> None:
    """
    对2个不同的工作薄执行vlookup操作

    主工作薄：需要写入数据的工作薄
    数据工作薄：根据模板工作薄提供的条件（列）在数据工作薄中查找，提供数据来源的工作薄
    注意！函数执行完后，只对wb_main对象进行了数据写入。在函数外部还需wb_main.save("filename.xlsx"),vlookup后的数据才能保存为excel表。

    20220714 test OK

    :param:
        'wb_main': 主工作薄对象
        'ws_main_index': 需要处理的主工作表索引号
        'main_key': 主工作表key所在列号
        'main_value': 主工作表value需要填写的列号
        'wb_source': 数据工作薄对象
        'ws_source_index': 需要处理的数据工作表索引号
        'source_key':  数据工作表key所在列号
        'source_value': 数据工作表value所在列号
        'line' :从第几行开始vlookup

    :type:
        'wb_main': class Workbook
        'ws_main_index': int
        'main_key': str
        'main_value': str
        'wb_source': class Workbook
        'ws_source_index': int
        'source_key':  str
        'source_value': str
        'line' :int

    :return: None

    """
    # 获取数据工作表对象
    ws_source = wb_source[wb_source.sheetnames[ws_source_index-1]]
    # 获取主工作表对象
    ws_main = wb_main[wb_main.sheetnames[ws_main_index-1]]
    # 获取数据工作表的查找列和数据列，分别生成2个元组
    source_key_tuple = ws_source[source_key]
    source_value_tuple = ws_source[source_value]
    # 创建2个列表，遍历元组，将元组中每个单元格的值添加到列表中
    list_key =[]
    list_value = []
    for cell in source_key_tuple:
        list_key.append(cell.value)
    for cell in source_value_tuple:
        list_value.append(cell.value)
    # 通过数据工作表的key列和value列，创建好需要进行vlookup的字典
    dic = dict(zip(list_key,list_value))
    # 将列号 （str）转为列索引值 （int）
    main_key_index = column_index_from_string(main_key)
    main_value_index = column_index_from_string(main_value)
    # 从第line行开始进行vlookup  可根据表头行数进行修改
    for row in range(int(line),ws_main.max_row+1):
        # ------------------------------------------------------
        # 采用dict[key]方式查字典，如果没有key的话会raise KeyError
        # try:
        #     ws_main.cell(row=row,column=template_value_index).value = dic[
        #     ws_main.cell(row=row,column=template_key_index).value]
        # except KeyError:
        #     #找不到数据 相应的单元格填上#N/A
        #     ws_main.cell(row=row, column=template_value_index).value = "#N/A"
        # -------------------------------------------------------
        #采用dict.get()避免出现keyerror
        ws_main.cell(row=row, column=main_value_index).value = dic.get(
            ws_main.cell(row=row, column=main_key_index).value)


if __name__ == '__main__':
    # wbmain = openpyxl.load_workbook('tests\\tests_easy_vlookup\\主表.xlsx')
    # wbsource = openpyxl.load_workbook('tests\\tests_easy_vlookup\\数据表.xlsx')
    # easyvlookup(wb_main=wbmain,
    #             ws_main_index=1,
    #             main_key='B',
    #             main_value='D',
    #             wb_source=wbsource,
    #             ws_source_index=2,
    #             source_key='K',
    #             source_value='J',
    #             line=2)
    # wbmain.save('已进行vlookup.xlsx')
    wbmain = openpyxl.load_workbook('tests\\tests_easy_vlookup\\20220615_2021年度失业保险支持企业稳定岗位返名单（大型企业返还比例从30%转为50%）(添加银行信息).xlsx')
    wbsource = openpyxl.load_workbook('tests\\tests_easy_vlookup\\全市2021年度失业保险稳岗返还企业数据 20220418整理(市局王岳岳20220419提供).xlsx')
    easyvlookup(wb_main=wbmain,
                ws_main_index=1,
                main_key='D',
                main_value='R',
                wb_source=wbsource,
                ws_source_index=3,
                source_key='C',
                source_value='L',
                line=2)
    wbmain.save('已进行vlookup2.xlsx')

