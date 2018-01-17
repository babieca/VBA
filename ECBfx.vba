Sub save_FX_rates()
    Application.DecimalSeparator = "."
    Application.ThousandsSeparator = ","
    Application.UseSystemSeparators = True
        
    Dim ipddbb As String
    Dim nameddbb As String
    Dim portddbb As Integer
    Dim userddbb As String
    Dim passddbb As String
    Dim msg As String
    
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Dim sht As Worksheet
    
    ipddbb = ""
    portddbb = 3306
    nameddbb = ""
    userddbb = ""
    passddbb = ""
    
    'msg = validateFields
    'If msg <> "" Then
    '    MsgBox (msg)
    '    Exit Sub
    'End If
    
    Set conn = New ADODB.Connection
        
    conn.Open "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
                "SERVER=" & ipddbb & ";" & _
                "PORT=" & CStr(portddbb) & ";" & _
                "DATABASE=" & nameddbb & ";" & _
                "USER=" & userddbb & ";" & _
                "PASSWORD=" & passddbb & ";" & _
                "OPTION=3;"""
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    
    Set sht = ActiveSheet
    
    tblddbb = "lbv.histccy"
    
    lastCol = sht.Cells(1, Columns.Count).End(xlToLeft).Column
    lastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
    counter = 0
    If lastRow > 1 Then

        resp = MsgBox("Insert into MySQL:" & vbCrLf & _
                      "<Yes> Save" & vbCrLf & _
                      "<No> Abort" & vbCrLf, _
                      vbYesNo, _
                      "Assets to MySQL")
        
        If resp = 7 Then
            Exit Sub
        End If
        
        For j = 2 To lastRow
            dt = Format(Trim(sht.Cells(j, 1)), "yyyy-mm-dd")
            If dt > "2016-01-01" Then
                For i = 2 To lastCol
                    ccy = Trim(sht.Cells(1, i))
                    If ccy <> "" Then
                        xrate = sht.Cells(j, i)
                        If Not IsNumeric(xrate) Then
                            xrate = -1
                        End If
                        strSQL = "INSERT INTO " & tblddbb & _
                                    " ( ccy_name, ccy_date, ccy_xrate) "
                        strSQL = strSQL & " VALUES ('" & ccy & "','" & dt & "'," & xrate & ")"
                                    
                        strSQL = strSQL & " ON DUPLICATE KEY UPDATE " & _
                                    "ccy_name = '" & ccy & "', " & _
                                    "ccy_date = '" & dt & "', " & _
                                    "ccy_xrate = " & xrate & ";"
                        cmd.CommandText = strSQL
                        cmd.Execute strSQL
                        counter = counter + 1
                    End If
                Next i
            End If
        Next j
    End If
    MsgBox ("A total of " & CStr(counter) & " new records successfully inserted into MySQL.")
    conn.Close
    Set conn = Nothing
End Sub


