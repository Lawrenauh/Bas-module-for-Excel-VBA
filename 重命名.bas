Attribute VB_Name = "重命名"
Sub rename()
Attribute rename.VB_ProcData.VB_Invoke_Func = " \n14"

Dim s, str As String
    
    s = Application.InputBox("请输入完整路径：" & Chr(13), Type:=1 + 2)
    s = IIf(Right(s, 1) = "\", s, s & "\")
    str = Dir(s & "*")
    
    Do While str <> ""
        Name s & str As s & "市场推广费用计提 - " & str
        str = Dir
    Loop

End Sub
