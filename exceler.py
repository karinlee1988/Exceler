#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2021/8/1 15:12
# @Author : karinlee
# @FileName : exceler.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976
# @github : https://github.com/karinlee1988/
# @gitee : https://gitee.com/karinlee/
# @Personal website : https://karinlee.cn/

import os
import openpyxl

from openpyxl.utils import column_index_from_string  # ,get_column_letter


def get_singlefolder_fullfilename(folder_path:str,filetype:list) -> list:
    """
    获取待处理文件夹里指定后缀的文件名（单个文件夹，不包括子文件夹的文件）

    :param folder_path: 文件夹路径
    :type  folder_path: str

    :param filetype: 指定一种或多种类型文件的后缀列表（如['.xlsx']或['.xlsx','.xls']）
    :type  filetype: list

    :return: 文件夹里面所有全文件名列表（包含子文件夹里面的文件）
    :rtype:  list
    """
    filename_list = []
    files = os.listdir(folder_path)
    for file in files:
        # os.path.splitext():分离文件名与扩展名
        if os.path.splitext(file)[1] in filetype:
            filename_list.append(folder_path+'\\'+file)
    return filename_list

def vlookup(
        wb_template,
        ws_template_index:int,
        template_key:str,
        template_value:str,
        wb_source,
        ws_source_index:int,
        source_key:str,
        source_value:str,
        line:int
        ) -> None:
    """
    对2个不同的工作薄执行vlookup操作

    模板工作薄：需要写入数据的工作薄
    数据工作薄：根据模板工作薄提供的条件（列）在数据工作薄中查找，提供数据来源的工作薄
    注意！函数执行完后，只对wb_template对象进行了数据写入。在函数外部还需wb_template.save("filename.xlsx"),vlookup后的数据才能保存为excel表。

    :param:
        'wb_template': 模板工作薄对象
        'ws_template_index': 需要处理的模板工作表索引号
        'template_key': 模板工作表key所在列号
        'template_value': 模板工作表value需要填写的列号
        'wb_source': 数据工作薄对象
        'ws_source_index': 需要处理的数据工作表索引号
        'source_key':  数据工作表key所在列号
        'source_value': 数据工作表value所在列号
        'line' :从第几行开始vlookup

    :type:
        'wb_template': class Workbook
        'ws_template_index': int
        'template_key': str
        'template_value': str
        'wb_source': class Workbook
        'ws_source_index': int
        'source_key':  str
        'source_value': str
        'line' :int

    :return: None

    """
    # 获取数据工作表对象
    ws_source = wb_source[wb_source.sheetnames[ws_source_index]]
    # 获取模板工作表对象
    ws_template = wb_template[wb_template.sheetnames[ws_template_index]]
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
    template_key_index = column_index_from_string(template_key)
    template_value_index = column_index_from_string(template_value)
    # 从第line行开始进行vlookup  可根据表头行数进行修改
    for row in range(int(line),ws_template.max_row+1):
        # ------------------------------------------------------
        # 采用dict[key]方式查字典，如果没有key的话会raise KeyError
        # try:
        #     ws_template.cell(row=row,column=template_value_index).value = dic[
        #     ws_template.cell(row=row,column=template_key_index).value]
        # except KeyError:
        #     #找不到数据 相应的单元格填上#N/A
        #     ws_template.cell(row=row, column=template_value_index).value = "#N/A"
        # -------------------------------------------------------
        #采用dict.get()避免出现keyerror
        ws_template.cell(row=row, column=template_value_index).value = dic.get(
            ws_template.cell(row=row, column=template_key_index).value)

def worksheet_save_as(workbook) -> None:
    """
    将一个工作薄里面的多个工作表分别另存为独立的工作薄，独立的工作薄名称为原工作薄各工作表表名

    :param workbook:需要进行工作表另存为的workbook对象
    :type workbook: class Workbook

    :return: None

    """
    sheetname_list = workbook.sheetnames
    for name in sheetname_list:
        worksheet = workbook[name]
        # 创建新的Excel
        workbook_new = openpyxl.Workbook()
        # 获取当前sheet
        worksheet_new = workbook_new.active
        # 两个for循环遍历整个excel的单元格内容
        for i, row in enumerate(worksheet.iter_rows(),start=1): #enumerate()从1开始  # 或者for i, row in enumerate(worksheet.rows):
            for j, cell in enumerate(row,start=1):#enumerate()从1开始
                # 写入新Excel
                worksheet_new.cell(row=i, column=j, value=cell.value)
                # 设置新Sheet的名称
                worksheet_new.title = name
        workbook_new.save(name + '.xlsx')

def merge_xlsx_workbooks(folder_path:str, sheet_index:str or int, title_row:str or int ,solid_column:str) -> None:
    """
    将某个文件夹下面所有的.xlsx文件中某个sheet页的内容合并为一个工作薄
    仅支持.xlsx格式
    合并后的工作薄第一列的内容会指明该行内容来源为哪个.xlsx文件

    20210801 test OK

    :param folder_path: 待合并工作薄文件夹路径
    :type folder_path: str

    :param sheet_index: 待合并的sheet页(从1开始)
    :type sheet_index: str or int

    :param title_row: 表头行数（从1开始）
    :type title_row: str or int

    :param solid_column: 按照某列存在的数据进行合并，避免合并空行
    :type solid_column: str

    :return: None
    :rtype:  None
    """

    # 转换为int，确保后续处理没问题
    counter = 0
    sheet_index = int(sheet_index)
    title_row = int(title_row)
    # 获取path路径下所有xlsx文件的文件名（文件名包含路径）
    xlsx_filename_list = get_singlefolder_fullfilename(folder_path=folder_path,filetype=['.xlsx'])
    # 新建合并后的工作薄
    main_workbook = openpyxl.Workbook()
    main_worksheet = main_workbook.active
    # 构建表头 表头数据随便从文件列表中第一个文件处获取
    temp_workbook = openpyxl.load_workbook(xlsx_filename_list[0])
    temp_worksheet = temp_workbook[temp_workbook.sheetnames[sheet_index - 1]]
    list_all_title = []
    list_row_title = []
    for each_title_row in range(1,title_row+1):
        for cell in temp_worksheet[each_title_row]:
            list_row_title.append(cell.value)
        list_all_title.append(list_row_title)
        list_row_title = []
    for title in list_all_title:
        # 在最左边插入空列，便于后续处理
        title.insert(0,'')
        main_worksheet.append(title)
    # 构建数据部分，需要从每个文件中拿取
    for filename in xlsx_filename_list:

        merge_workbook = openpyxl.load_workbook(filename)
        merge_worksheet = merge_workbook[merge_workbook.sheetnames[sheet_index - 1]]
        list_all = []
        list_row = []
        for row in range(title_row+1, merge_worksheet.max_row + 1):
            # 合并工作薄的表格中，某些行可能会有空数据，使用solid_column变量确保某列存在数据的才被合并进去
            #-------------------------------------------------------------------------------------
            # 这里也可以改为其他条件，确保符合条件的数据行才会被提取并写入到合并后的工作薄
            if merge_worksheet.cell(row=row,column=column_index_from_string(solid_column)).value:
            #-------------------------------------------------------------------------------------
                for cell in merge_worksheet[row]:  # 对当前行遍历所有单元格
                    list_row.append(cell.value)  # list_row是临时的1维列表 在遍历单元格获得每个单元格的值后写入列表 从而存储当前行的数据
                list_all.append(list_row)  # list_all是二维列表 里面的每个元素都是1个list_row
                list_row = []  # 重新初始化list_row列表
        for row in list_all:  # 对于2维列表的每个元素（每个元素就是每个1维列表 这些1维列表就等于excel表中1行的数据
            # 在最左边插入一列，内容为合并前的文件来源
            row.insert(0,os.path.basename(filename))

            main_worksheet.append(row)  # 可以直接使用worksheet.append()方法写入工作表中
        counter += 1
        print(f">>>{counter}<<<  {os.path.basename(filename)}  已写入合并工作薄中！")

    main_workbook.save("合并完成.xlsx")



if __name__ == '__main__':
   merge_xlsx_workbooks("各单位汇总\\",1,3,'C')
