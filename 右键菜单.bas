Attribute VB_Name = "�Ҽ��˵�"
Sub �г��Ҽ��˵�()

On Error Resume Next

Dim mc As CommandBarControl
    x = 1
    Cells(x, 1) = "Index"
    Cells(x, 2) = "Caption"
    Cells(x, 3) = "ID"
    
    For Each mc In Application.CommandBars("cell").Controls
        Cells(x, 1) = mc.Index
        Cells(x, 2) = mc.Caption
        Cells(x, 3) = mc.ID
        x = x + 1
    Next

End Sub

Sub ɾ���Ҽ��˵���()

    Application.CommandBars("cell").Controls(18).delete

End Sub
