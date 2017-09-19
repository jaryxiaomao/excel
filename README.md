# excel使用的一个小技巧
使用 宏 删除指定的列

`s = Array("xiaomao", 1)
    For i = [iv1].End(xlToLeft).Column To 1 Step -1
        For Each c In s
            If Cells(1, i) = c Then Cells(1, i).EntireColumn.Delete
        Next
    Next`
    
    
    
   其中的ivl代表选中的活动表的第一行，如果行数为变量则可以表示为Range("iv" & a)，Rang.End属性返回一个Rang对象，该对象代表包含源区域的区域结尾处的<br>元格。End(xlToLeft)相当于相当于在源区域按Ctrl+左方向键。xlToRight、xlUp、xlDown为其他三个方向。
    
    End（xlUp）：若活动单元格为空，其上一个单元格也为空，将会向上寻找该列第一次出现的非空单元格；
                若活动单元格非空，其上一个单元格也非空，将会选中活动单元格所在列的最后一个非空单元格；
     
     
     
     
例子： for a to b step c
      默认step为1，step为c时遍历相邻数的间隔为c，从尾部遍历step为负数
cells（）表示单元格
