Attribute VB_Name = "右键菜单"
Sub 列出右键菜单()

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

Sub 删除右键菜单项()

    Application.CommandBars("cell").Controls(18).delete

End Sub
