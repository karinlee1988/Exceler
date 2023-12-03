# Exceler

用于实现excel 文档处理的工具库



### 特性

该库用于处理EXCEL表格，如表格的拆分、合并、数据匹配等功能。

请注意，该库是用于处理EXCEL表格/文档，而不单是处理数据。在处理EXCEL表格/文档的过程中，尽量保证处理后与处理前表格的样式一致。因此，本库大部分使用openpyxl库而不是更方便的pandas库来进行处理。

限于该库开发的目的和编者个人技术，部份功能运行占用资源和耗时较长。



### 开始使用
每项功能均建立单独的.py文件，若要使用某项功能，则直接运行该.py文件即可。

部分.py里有图形界面类 `class XXXGUI（object）` 提供图形界面。

独立的VBA文件夹里包含了一些excel VBA 工具。

---



#### worksheet_split_by_column.py

**说明：**    用于按列的内容拆分excel .XLSX工作薄中的一个工作表，分别另存为多个独立的工作薄。打印可选。

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

**说明：**    功能类似excel中的vlookup函数，用于多表数据匹配对碰使用。一般情况下运行速度比excel vlookup函数要快。

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
    wbmain = openpyxl.load_workbook('tests\\test_easy_vlookup\\主表.xlsx')
    wbsource = openpyxl.load_workbook('tests\\test_easy_vlookup\\数据表.xlsx')
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

**easy_vlookup.py**     提供`EasyVlookupGui(object)`图形界面类，通过`app = EasyVlookupGui()`调用运行图形界面。

---


#### worksheets_save_as.py

**说明：**    将一个工作薄里面的多个工作表分别另存为独立的工作薄，独立的工作薄名称为原工作薄各工作表表名。

**结构：**

```python
def worksheet_save_as(workbook_fullfilename:str) -> None:
    """
    将一个工作薄里面的多个工作表分别另存为独立的工作薄，独立的工作薄名称为原工作薄各工作表表名
    注意：拆分后表格格式可能有所变化

    20220928 test OK

    :param workbook_fullfilename:需要进行工作表另存为的.xlsx文件全文件名
    :type workbook_fullfilename: str

    :return: None

    """
```

**调用示例：**

```python
if __name__ == '__main__':
    wb = 'tests\\test_worksheets_save_as\\拆分工作表测试.xlsx'
    worksheet_save_as(wb)
```



---

#### worksheet_split_by_regular_line.py

**说明：**      按照固定行数将单个工作表拆分为多个工作表(如将1个3000行的表格拆分为3个1000行的表格)，表头默认为1行。对于工作薄，操作对象默认为工作薄的第一个工作表。

**结构：**

```python
class WorksheetSplitByRegularLine(object):
    """
    按照固定行数将单个工作表拆分为多个工作表(如将1个3000行的表格拆分为3个1000行的表格)
    表头默认为1行
    对于工作薄，操作对象默认为工作薄的第一个工作表

    20221002 test OK
    """

    def __init__(self,full_filename:str,batche_num:int or str):
        """

        :param full_filename: 待拆分xlsx文件的全文件名（相对路径或绝对路径）
        :param batche_num: 需要拆分的固定行数

        :type full_filename: str
        :type batche_num: int or str
        """
```

**调用示例：**

```python
if __name__ == '__main__':
	cut = WorksheetSplitByRegularLine("tests\\test_worksheet_split_by_regular_line\\固定行数拆分测试.xlsx",100)
    cut.main()
```



---
#### merge_worksheets.py

**说明：**      将某个文件夹下面所有的.xlsx文件中某个sheet页的内容合并,生成一个工作薄，仅支持.xlsx格式。合并后的工作薄第一列的内容会指明该行内容来源为哪个.xlsx文件。程序运行后，生成 ‘合并完成.xlsx’的合并文件。

**结构：**

- MergeXlsxWorkSheets(object)类

```python
class MergeXlsxWorkSheets(object):
    """
    将某个文件夹下面所有的.xlsx文件中某个sheet页的内容合并,生成一个工作薄
        仅支持.xlsx格式
        合并后的工作薄第一列的内容会指明该行内容来源为哪个.xlsx文件

    20221004 test OK

    """

    def __init__(self,folder_path:str):
        """
        :param folder_path: 待合并工作薄文件夹路径
        :type folder_path: str
        """
        self.folder_path = folder_path
```

- MergeXlsxWorkSheets(object)类中的merge_xlsx_workbooks(self,...) 实例方法

```python
    def merge_xlsx_workbooks(self,
                             sheet_index:str or int, 
                             title_row:str or int ,
                             solid_column:str) -> None:
        """
        合并主程序入口

        :param sheet_index: 待合并的sheet页(从1开始)
        :type sheet_index: str or int

        :param title_row: 表头行数（从1开始）
        :type title_row: str or int

        :param solid_column: 按照某列存在的数据进行合并，避免合并空行(列号)
        :type solid_column: str

        :return: None
        :rtype:  None
        """
```


**调用示例：**

先传入待合并xlsx文件的文件夹路径，实例化后传入工作表序号（从1开始）、表头行数（从1开始）、检测列号（检测是否为空行）后运行合并，生成 ‘合并完成.xlsx’的合并文件。

```python
if __name__ == '__main__':
    merger = MergeXlsxWorkSheets('tests\\test_merge_worksheets')
    merger.merge_xlsx_workbooks(sheet_index=1,title_row=3,solid_column='C')

```
---

### VBA文件夹

####  将xls文件批量转为xlsx文件

该文件夹包含了可以将.xls文件批量转为.xlsx文件的工具：`Convert2xlsx.xlsm`

打开后按里面的指引运行宏即可。

