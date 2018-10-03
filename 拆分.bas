Attribute VB_Name = "���"
Sub ��ֹ�����()
Attribute ��ֹ�����.VB_ProcData.VB_Invoke_Func = " \n14"

Dim str As String
Dim dic
Dim rng, cell As Range

    Set dic = CreateObject("Scripting.Dictionary")
    str = ActiveWorkbook.path
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next

Line1:

    Set rng = Application.InputBox(prompt:="��ѡ��Ҫ��ֵ��У�" & Chr(13), Type:=8)
    If IsEmpty(rng) Or rng Is Nothing Then
        Exit Sub
    End If
    
    Do While rng.Columns.Count > 1
        MsgBox "��ʾ��ѡ�����򳬹�һ�У�������ѡ��"
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
    
        
        '����Ӧ�п�
        'ActiveSheet.UsedRange.EntireColumn.AutoFit

        If Dir(str & "\���", vbDirectory) = "" Then
            MkDir str & "\���"
        End If
        
        ActiveWorkbook.SaveAs Filename:=str & "\���\չʾ�г��ƹ���ü���_" & item & "_201801"
        
        'ActiveWorkbook.SaveAs Filename:=str & "\���\�����λ����_" & item
        'ActiveWorkbook.SaveAs Filename:=str & "\���\�ɱ�¼��_" & item, FileFormat:=xlCSV
        'ActiveWorkbook.SaveAs Filename:=str & "\���\չʾ��������2017.2-7.24��- " & Item & ".xlsx"
        
        ActiveWorkbook.Save
        ActiveWorkbook.Close
        
    Next
    
    ActiveSheet.AutoFilterMode = False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub
