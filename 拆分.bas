Attribute VB_Name = "拆分"
Sub 拆分工作表()
Attribute 拆分工作表.VB_ProcData.VB_Invoke_Func = " \n14"

Dim str As String
Dim dic
Dim rng, cell As Range

    Set dic = CreateObject("Scripting.Dictionary")
    str = ActiveWorkbook.path
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next

Line1:

    Set rng = Application.InputBox(prompt:="请选择要拆分的列：" & Chr(13), Type:=8)
    If IsEmpty(rng) Or rng Is Nothing Then
        Exit Sub
    End If
    
    Do While rng.Columns.Count > 1
        MsgBox "提示：选择区域超过一列，请重新选择！"
        Set rng = Nothing
        GoTo Line1
    Loop
    
    For Each cell In Range(rng(2), rng(rng.Count).End(xlUp))
        dic(cell.value) = 1
    Next
    
    For Each item In dic
    
        ActiveSheet.AutoFilterMode = False
        ActiveSheet.UsedRange.AutoFilter Field:=rng.Column, Criteria1:=item
        ActiveSheet.UsedRange.Copy
        Workbooks.Add
        
        With selection
            .PasteSpecial Paste:=xlPasteColumnWidths
            .PasteSpecial Paste:=xlPasteValues
            .PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        End With
        
        ActiveWindow.SplitRow = 1
        ActiveWindow.FreezePanes = True
        Application.CutCopyMode = False
    
        
        '自适应列宽
        'ActiveSheet.UsedRange.EntireColumn.AutoFit

        If Dir(str & "\拆分", vbDirectory) = "" Then
            MkDir str & "\拆分"
        End If
        
        ActiveWorkbook.SaveAs Filename:=str & "\拆分\展示市场推广费用计提_" & item & "_201801"
        
        'ActiveWorkbook.SaveAs Filename:=str & "\拆分\错误点位详情_" & item
        'ActiveWorkbook.SaveAs Filename:=str & "\拆分\成本录入_" & item, FileFormat:=xlCSV
        'ActiveWorkbook.SaveAs Filename:=str & "\拆分\展示广告促销（2017.2-7.24）- " & Item & ".xlsx"
        
        ActiveWorkbook.Save
        ActiveWorkbook.Close
        
    Next
    
    ActiveSheet.AutoFilterMode = False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub
