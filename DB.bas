Attribute VB_Name = "DB"
Sub ConnectMySQL()
Attribute ConnectMySQL.VB_ProcData.VB_Invoke_Func = "J\n14"

Dim r As Integer
Dim sql As String
Dim con As ADODB.Connection

    Application.ScreenUpdating = False

    Set con = New ADODB.Connection
    Set Rec = New Recordset
    
    
    con.CommandTimeout = 720
    con.ConnectionString = "Driver={MySql ODBC 5.3 Unicode Driver};" + "Server=localhost;" + "DB=sales_promotion;" + "UID=root;" + "PWD=031226;" + "OPTION=3;"
    
    con.Open
    
    'If con.State = adStateOpen Then
    '    MsgBox "����״̬��" & con.State & vbCrLf & "ADO�汾��" & con.Version, vbInformation, ""
    'End If
    
    'con.Execute "update user set authentication_string = password('HH031226') where user='huhuan3';", , adCmdText
    'con.Execute "flush privileges;"
    
    sql = "select s1.ad_date 'ʱ��', s3.pid '��ĿID', s3.tagid '���λID', sum(imp) '�����ع�', sum(uimp) '�����ع�', sum(clk) '������', sum(uclk) '�������', sum(estImp) 'Ԥ���ع�', sum(estClk) 'Ԥ�����' from ad_data s1 left outer join ad_est s2 on concat_ws('-', s1.ad_date, s1.ad_id, s1.ad_tag_id) = concat_ws('-', s2.ad_date, s2.ad_id, s2.ad_tag_id) left outer join id_mapping s3 on concat_ws('-', s1.ad_id, s1.ad_media_id, s1.ad_tag_id) = concat_ws('-', s3.ad_id, s3.ad_media_id, s3.ad_tag_id) group by s1.ad_date, s3.pid, s3.tagid;"
    
    Set Rec = con.Execute(sql, , adCmdText)
    ActiveSheet.Range("a1:j1").value = Array("����", "ʱ��", "��ĿID", "���λID", "�����ع�", "�����ع�", "������", "�������", "Ԥ���ع�", "Ԥ�����")
    ActiveSheet.Range("b2").CopyFromRecordset Rec
    
    con.Close: Set con = Nothing
    
    r = Range("a1").CurrentRegion.Rows.Count
    ActiveSheet.Range("A2:A" & r).FormulaR1C1 = "=RC[1]&""-""&RC[2]&""-""&RC[3]"
    
    Columns("A:K").AutoFit
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    Application.ScreenUpdating = True

End Sub
