Attribute VB_Name = "������"
Sub rename()
Attribute rename.VB_ProcData.VB_Invoke_Func = " \n14"

Dim s, str As String
    
    s = Application.InputBox("����������·����" & Chr(13), Type:=1 + 2)
    s = IIf(Right(s, 1) = "\", s, s & "\")
    str = Dir(s & "*")
    
    Do While str <> ""
        Name s & str As s & "�г��ƹ���ü��� - " & str
        str = Dir
    Loop

End Sub
