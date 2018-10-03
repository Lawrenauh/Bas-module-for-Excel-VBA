Attribute VB_Name = "常用"
Sub 格式化单元格()
Attribute 格式化单元格.VB_ProcData.VB_Invoke_Func = " \n14"

    selection.ClearFormats
    
End Sub

Sub 可见单元格()
Attribute 可见单元格.VB_ProcData.VB_Invoke_Func = "S\n14"

    selection.SpecialCells(xlCellTypeVisible).Select
    
End Sub

Sub 空格()
Attribute 空格.VB_ProcData.VB_Invoke_Func = "K\n14"

    Application.ScreenUpdating = False
    
    On Error GoTo Error1
    
    selection.SpecialCells(xlCellTypeBlanks).Select
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
Error1:
    MsgBox "警告：所选区域中不包含空单元格！", 48
    Resume Next
    
End Sub

Sub 公式转化成数值()
Attribute 公式转化成数值.VB_ProcData.VB_Invoke_Func = "N\n14"

    Application.ScreenUpdating = False
    
    Set M_cells = Intersect(selection.SpecialCells(xlCellTypeVisible), selection)    '注意：当执行Selection.SpecialCells()命令时，若Selection为工作表中的单个单元格，此时Selection默认为整张工作表！
    
    For Each rng In M_cells
    
        If IsNumeric(rng) Then
        
            If Len(rng) > 15 Then
                rng.NumberFormatLocal = "@"
            End If
            
        End If
        
        rng.value = rng.value
     
    Next
     
    Application.ScreenUpdating = True
     
End Sub

Sub 一般格式()
Attribute 一般格式.VB_ProcData.VB_Invoke_Func = "Y\n14"

    Application.ScreenUpdating = False

    With selection
    
        .Font.Bold = False
        .Font.Italic = False
        .Font.Underline = xlUnderlineStyleNone
        .Font.Size = 9
        
        .Borders.LineStyle = xlContinuous
        
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
         
        .RowHeight = 21
        .ColumnWidth = 9
        
    End With
    
    Application.ScreenUpdating = True
    
End Sub


Sub 复制粘贴到不连续区域()
Attribute 复制粘贴到不连续区域.VB_ProcData.VB_Invoke_Func = "V\n14"

Dim MyData As New DataObject

Application.ScreenUpdating = False


If Not Application.CutCopyMode = False Then

   MyData.GetFromClipboard
   
   Data1 = MyData.GetText(1)
   
   Data2 = Replace(Data1, Chr(9), "_")
   Data3 = Replace(Data2, Chr(13) + Chr(10), "_")

   arr = Split(Data3, "_")
   

   If selection.Rows.Count = 1 Or selection.Columns.Count = 1 Then
   
   Else
   
      MsgBox "目标区域必须是一行或一列！请重新选择。", 48, "警告！"
      Exit Sub
      
   End If
   
   
   n = 0
   j = UBound(arr)
   
      
   If arr(j) = Chr(32) Then
      j = j - 1
   End If

   
   For Each rng In selection.SpecialCells(xlCellTypeVisible)
   
       i = n Mod j
       rng.value = arr(i)
       n = n + 1
       
   Next
   
      
   Set MyData = Nothing
   
Else

   MsgBox "剪切板为空！", 48, "警告！"
   
End If

  
Application.ScreenUpdating = True

End Sub

Sub 批量删除工作表()

Application.ScreenUpdating = False

MyPath = ActiveWorkbook.path
MyName = Dir(MyPath & "\" & "*.xls")

s = Application.InputBox("请输入要删除的工作表名称：" & Chr(13), Type:=1 + 2)

Do While MyName <> ""

    Workbooks.Open(MyPath & "\" & MyName).Activate
    
    If Sheets.Count <> 1 Then
        ActiveWorkbook.Worksheets(s).delete
        Application.DisplayAlerts = False
        ActiveWorkbook.Close True
    End If

    MyName = Dir

Loop

Application.Quit

Application.ScreenUpdating = True

End Sub

Sub 行选择至最下面单元格()
Attribute 行选择至最下面单元格.VB_ProcData.VB_Invoke_Func = "X\n14"

    Dim area As Range
    
    Set area = ActiveCell

    Application.ScreenUpdating = False
    
    For Each rng In selection
    
        r = rng.SpecialCells(xlLastCell).Row
        c = rng.Column

        Set area = Union(area, Union(rng, Range(rng, Cells(r, c))))
        
    Next
    
    Application.ScreenUpdating = True
    
    Cells(r, c).Activate
    
    Application.ScreenUpdating = False
    
    area.Select
    
    Application.ScreenUpdating = True
        
End Sub

Sub 合并单元格()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim rng As Range

    a = selection.Row
    b = selection.Column
    n = selection.Rows.Count
    
    Add = Application.InputBox(prompt:="请输入要合并单元格的频数：" & Chr(13), Type:=0 + 1)
    
    If Add = 0 Then
        Exit Sub
    End If
    
    selection.MergeCells = False
    
    For i = a To a + n - 1
    
        If m = 0 Then
            Cells(i, b).Select
        End If
        
        m = m + 1
        
        Union(selection, Cells(i, b)).Select
        
        If (i - a + 1) Mod Add = 0 Then
            selection.Merge
            m = 0
        End If
        
    Next
    
Application.DisplayAlerts = True
Application.ScreenUpdating = True
    
End Sub

Sub 跳过空单元格粘贴()
Attribute 跳过空单元格粘贴.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False

    selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
    
    Application.ScreenUpdating = True
    
End Sub
        
Sub 格式整理()
Attribute 格式整理.VB_ProcData.VB_Invoke_Func = "F\n14"

    Application.ScreenUpdating = False
    
    Dim crr()

    t = Timer

    arr = Array(11, 16, 82, 94, 108)
    brr = Array(23, 27, 63, 64, 76, 77, 91, 92, 98, 105, 106, 115, 119, 137)
    
    x = False
    y = False
    
    a = selection(1).Row
    b = selection(1).Column
    
    n = selection.Rows.Count
    m = selection.Columns.Count
    
    ReDim crr(1 To m)
    
    For l = 1 To m
        crr(l) = Cells(1, b + l - 1).Interior.ColorIndex
    Next
    
    selection.ClearFormats
    
    For l = 1 To m
        Cells(1, b + l - 1).Interior.ColorIndex = crr(l)
    Next
    
    t1 = Timer
    
    For i = b To b + m - 1
    
        For Each a In arr
            If a = i Then
                x = True
                Exit For
            Else
                For Each b In brr
                    If b = i Then
                        y = True
                        Exit For
                    End If
                Next
            End If
        Next
        
        
        If x Then
            
            Columns(i).NumberFormatLocal = "@"
            
        Else
            If y Then
                Columns(i).NumberFormatLocal = "yyyy/m/d"
            Else
                If i = 28 Then
                    Columns(i).NumberFormatLocal = "[$-F400]h:mm:ss AM/PM"
                Else
                    Columns(i).NumberFormatLocal = "G/通用格式"
                End If
            End If
        End If
        
        x = False
        y = False
    
    Next
    
    t2 = Timer
    
    For Each cell In selection
    
        If cell.value = "" Then
            cell.ClearContents
        Else
            cell.value = cell.value
        End If

    Next
    
    t3 = Timer
    
    With selection
    
        .Font.Bold = False
        .Font.Italic = False
        .Font.Underline = xlUnderlineStyleNone
        .Font.Size = 9
        
        .Borders.LineStyle = xlContinuous
        
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
         
        .RowHeight = 21
        .ColumnWidth = 9
        
    End With
    
    t4 = Timer
    
    If b <= 23 And b + m - 1 >= 23 Then
    
        For Each cell In Range(Cells(1, 23), Cells(n, 23))
            If cell Like "1905/" & "*" Then
                cell.NumberFormatLocal = "G/通用格式"
            End If
        Next
        
    End If
    
    t5 = Timer
    
    MsgBox "总共运行时间：" & Timer - t & "s" '& Chr(10) & "其中：" & Chr(10) & "格式化：" & t1 - t & "s" & Chr(10) & "每行格式设定：" & t2 - t1 & "s" & Chr(10) _
            & "单元格处理：" & t3 - t2 & "s" & Chr(10) & "边框字体行距：" & t4 - t3 & "s" & Chr(10) & "婚礼日期整理：" & t5 - t4 & "s"
    
    Application.ScreenUpdating = True

End Sub

Sub 选中列()
Attribute 选中列.VB_ProcData.VB_Invoke_Func = "C\n14"

Application.ScreenUpdating = False

Set area = ActiveCell

    For Each cell In selection
        Set area = Union(area, Columns(cell.Column))
    Next
    
    area.Select
Application.ScreenUpdating = True

End Sub

Sub 选中行()
Attribute 选中行.VB_ProcData.VB_Invoke_Func = "r\n14"

Application.ScreenUpdating = False

Set area = ActiveCell

    For Each cell In selection
        Set area = Union(area, Rows(cell.Row))
    Next
    
    area.Select
    
Application.ScreenUpdating = True
    
End Sub

Sub 字符添加与删减()
Attribute 字符添加与删减.VB_ProcData.VB_Invoke_Func = "m\n14"

Application.ScreenUpdating = False

    Do Until u = "+" Or u = "-"
        u = InputBox("请选择处理方式：" & Chr(10) & Chr(10) & "     " & "+：添加字符" & "         " & "-：删除字符", "提示")
        If u = "" Then
            Exit Sub
        Else
            If u <> "+" And u <> "-" Then
                MsgBox " 抱歉，请输入规定的字符 + 或 - ！", vbInformation, "提示"
            End If
        End If
    Loop
        
    Do Until v = 1 Or v = 2
        v = InputBox("请选择位置：" & Chr(10) & Chr(10) & "     " & "1：首位" & "   " & "2：末位", "提示")
        If v = "" Then
            Exit Sub
        Else
            If v <> 1 And v <> 2 Then
                MsgBox " 抱歉，请输入规定的数值 1 或 2 ！", vbInformation, "提示"
            End If
        End If
    Loop
        
    w = InputBox("请输入字符：", "提示")
        
    For Each rng In selection.SpecialCells(xlCellTypeVisible)
    
        If u = "+" Then
            If v = 1 Then
                rng.value = w & rng
            Else
                rng.value = rng & w
            End If
        Else
            If v = 1 Then
                If Left(rng, Len(w)) = w Then
                    rng.value = Right(rng, Len(rng) - Len(w))
                End If
            Else
                If Right(rng, Len(w)) = w Then
                    rng.value = Left(rng, Len(rng) - Len(w))
                End If
            End If
        End If
    
    Next

Application.ScreenUpdating = True

End Sub

Sub 向下追加指定数目单元格()
Attribute 向下追加指定数目单元格.VB_ProcData.VB_Invoke_Func = "P\n14"

Dim c, r, Add, Num As Integer

c = ActiveCell.Column
r = ActiveCell.Row
Num = 0

Add = Application.InputBox(prompt:="请输入要追加选中的行数" & Chr(13), Type:=0 + 1)

If Add = False Then
    Exit Sub
End If


Do Until Num = Add - 1

   r = r + 1

   If Cells(r, c).EntireRow.Hidden Then
   
   Else
        Num = Num + 1
        
   End If

Loop


Range(ActiveCell, Cells(r, c)).SpecialCells(xlCellTypeVisible).Select

End Sub

Sub 隐藏灰色列()
Attribute 隐藏灰色列.VB_ProcData.VB_Invoke_Func = "H\n14"

    Application.ScreenUpdating = False

    k = Cells(1, 1).End(xlToRight).Column
    
    For j = 1 To k
    
        If Cells(1, j).Interior.ColorIndex = 15 Then
            Columns(j).Hidden = 1
        End If
               
    Next
    
    'Columns("CP").Hidden = 0
    'Columns("BR").Hidden = 0
    
    Application.ScreenUpdating = True

End Sub

Sub 去掉工作簿空单元格()

For Each Sheet In Worksheets

    'If Sheet.Name <> "简易索票" Then
        
        For Each cell In Sheet.UsedRange
        
            If cell.value = "" Then
                cell.ClearContents
            Else
                cell.value = cell.value
            End If
    
        Next
        
    'End If
    
Next

End Sub

Sub 条件格式隐藏单元格0值()

    Application.ScreenUpdating = False
    
    selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=0"
    selection.FormatConditions(selection.FormatConditions.Count).SetFirstPriority
    
    With selection.FormatConditions(1).Font
        .color = ActiveCell.Interior.color
    End With
    
    selection.FormatConditions(1).StopIfTrue = False
    
    Application.ScreenUpdating = True

End Sub

Sub 清除数据透视表无效数据项()
Attribute 清除数据透视表无效数据项.VB_ProcData.VB_Invoke_Func = "Q\n14"

'清除数据透视表下拉列表里显示无用的数据项

    Dim pvt As PivotTable, pvtcache As PivotCache
    Dim sht As Worksheet
    For Each sht In ActiveWorkbook.Worksheets
        For Each pvt In sht.PivotTables
            pvt.PivotCache.MissingItemsLimit = xlMissingItemsNone
        Next pvt
    Next sht
    On Error Resume Next
    For Each pvtcache In ActiveWorkbook.PivotCaches
        pvtcache.Refresh
    Next pvtcache
   On Error GoTo 0
End Sub

Sub 显示字段列表()
Attribute 显示字段列表.VB_ProcData.VB_Invoke_Func = "q\n14"

    If ActiveWorkbook.ShowPivotTableFieldList Then
        ActiveWorkbook.ShowPivotTableFieldList = False
    Else
        ActiveWorkbook.ShowPivotTableFieldList = True
    End If

End Sub

Sub 单元格颜色参数()

    c = ActiveCell.Interior.color
    
    H = Hex(c)
    
    s = 6 - Len(H)
    
    If s <> 0 Then
    
        For i = 1 To s
            H = "0" & H
        Next
        
    End If
    
    b1 = Left(H, 2)
    g1 = Mid(H, 3, 2)
    r1 = Right(H, 2)
    
    r2 = CLng("&h" & r1)
    g2 = CLng("&h" & g1)
    b2 = CLng("&h" & b1)
    
    MsgBox "背景色参数" & Chr(13) & Chr(13) & _
        "ColorIndex : " & ActiveCell.Interior.ColorIndex & Chr(13) & _
        "Color : " & ActiveCell.Interior.color & Chr(13) & _
        "RGB（" & r2 & " , " & g2 & " , " & b2 & "）", vbInformation, "单元格提示信息"
    
End Sub

Sub 颜色渐变()
Attribute 颜色渐变.VB_ProcData.VB_Invoke_Func = "B\n14"

    n = selection.Cells.Count

    color1 = color_rgb(selection(1))
    color2 = color_rgb(selection(n))
    
    i = 0
    
    For Each rng In selection
        k = i / n
        r = k * color1(0) + color2(0) * (1 - k)
        g = k * color1(1) + color2(1) * (1 - k)
        b = k * color1(2) + color2(2) * (1 - k)
        rng.Interior.color = RGB(r, g, b)
        i = i + 1
    Next

End Sub

Function color_rgb(rng As Range) As Integer()

Dim color(0 To 2) As Integer
    
    c = rng.Interior.color
    
    H = Hex(c)
    s = 6 - Len(H)
    
    If s <> 0 Then
        For i = 1 To s
            H = "0" & H
        Next
    End If
    
    b1 = Left(H, 2)
    g1 = Mid(H, 3, 2)
    r1 = Right(H, 2)
    
    color(0) = CLng("&h" & r1)
    color(1) = CLng("&h" & g1)
    color(2) = CLng("&h" & b1)
    
    color_rgb = color()

End Function
