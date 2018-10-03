Attribute VB_Name = "´ÙÏú"
Sub fuzhu()
Attribute fuzhu.VB_ProcData.VB_Invoke_Func = "F\n14"

Dim rng As Range

    Application.ScreenUpdating = False

    For Each rng In selection
        rng = CLng(rng.Offset(0, -3)) & "-" & rng.Offset(0, -2) & "-" & rng.Offset(0, -1)
    Next
    
    Application.ScreenUpdating = True

End Sub
