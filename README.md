# Exceler

用于实现excel 文档处理的工具库



### 特性

该库用于处理EXCEL表格，如表格的拆分、合并、数据匹配等功能。

请注意，该库是用于处理EXCEL表格/文档，而不单是处理数据。在处理EXCEL表格/文档的过程中，尽量保证处理后与处理前表格的样式一致。因此，本库大部分使用openpyxl库而不是更方便的pandas库来进行处理。

限于该库开发的目的和编者个人技术，部份功能运行占用资源和耗时较长。



### 开始使用
每项功能均建立单独的.py文件，若要使用某项功能，则直接运行该.py文件即可。



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

