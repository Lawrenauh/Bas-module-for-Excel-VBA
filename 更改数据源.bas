Attribute VB_Name = "更改数据源"
Sub Show_AllPivotDataSource()
    Dim i As Integer
    Dim pvt As PivotTable, pvtcache As PivotCache
    Dim sht As Worksheet
    
    i = 2
    For Each sht In ActiveWorkbook.Worksheets
        For Each pvt In sht.PivotTables
            ActiveSheet.Cells(i, 1) = sht.Name
            ActiveSheet.Cells(i, 2) = pvt.Name
            ActiveSheet.Cells(i, 3) = pvt.SourceData
            i = i + 1
        Next pvt
    Next sht
    On Error Resume Next
   On Error GoTo 0
End Sub

Sub Set_AllPivotSourceData()
    Dim pvt As PivotTable, pvtcache As PivotCache
    Dim sht As Worksheet
    
    For Each sht In ActiveWorkbook.Worksheets
        For Each pvt In sht.PivotTables
            pvt.SourceData = "Sheet1!$A:$U"
        Next pvt
    Next sht
    On Error Resume Next
   On Error GoTo 0
End Sub

