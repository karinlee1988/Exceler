# Exceler

用于实现excel 文档处理的工具库



### 特性

该库用于处理EXCEL表格，如表格的拆分、合并、数据匹配等功能。

请注意，该库是用于处理EXCEL表格/文档，而不单是处理数据。在处理EXCEL表格/文档的过程中，尽量保证处理后与处理前表格的样式一致。因此，本库大部分使用openpyxl库而不是更方便的pandas库来进行处理。

限于该库开发的目的和编者个人技术，部份功能运行占用资源和耗时较长。



### 开始使用
每项功能均建立单独的.py文件，若要使用某项功能，则直接运行该.py文件即可。

部分.py里有图形界面类 `class XXXGUI（object）` 提供图形界面。

---



#### worksheet_split_by_column.py

**说明：**用于按列的内容拆分excel .XLSX工作薄中的一个工作表，分别另存为多个独立的工作薄。打印可选。

**结构：**

```python
class WorksheetSplitByColumn(object):
    """
    按列拆分工作表为多个独立表格/拆分时打印可选
    20220712 test OK
    """


    def __init__(self,
                 filepath:str,
                 column_index:int,
                 title_index:int,
                 isprint:bool=False):

        """
        :param filepath: excel xlsx文件路径
        :param column_index: 按哪列的内容进行拆分（A列为1，B列为2,C列为3...）
        :param title_index: 表头的行数（从1开始，即表头有几行就填写几行）
        :param isprint: 是否拆分后自动打印（默认为False,即不启用打印功能）
        """
    ......
```



**调用示例：**

```python
if __name__ == '__main__':
    workbook_path = 'tests\\test_worksheet_split_by_column\\按列拆分测试表.xlsx'
    instance = WorksheetSplitByColumn(filepath=workbook_path,column_index=4,title_index=4,isprint=False)
    instance.main()
```

---

#### easy_vlookup.py

**说明：**功能类似excel中的vlookup函数，用于多表数据匹配对碰使用。一般情况下运行速度比excel vlookup函数要快。

**结构：**

```python
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
```

**调用示例：**

```python
if __name__ == '__main__':
    # 以下调用easyvlookup()函数 例1
    wbmain = openpyxl.load_workbook('tests\\tests_easy_vlookup\\主表.xlsx')
    wbsource = openpyxl.load_workbook('tests\\tests_easy_vlookup\\数据表.xlsx')
    easyvlookup(wb_main=wbmain,
                ws_main_index=1,
                main_key='B',
                main_value='D',
                wb_source=wbsource,
                ws_source_index=2,
                source_key='K',
                source_value='J',
                line=2)
    wbmain.save('已进行vlookup.xlsx')
```

**easy_vlookup.py** 提供`EasyVlookupGui(object)`图形界面类，通过`app = EasyVlookupGui()`调用运行图形界面。