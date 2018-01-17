'
'Sub changeDateFormats()
'    Set ActSheet = ActiveSheet
'    Set SelRange = Selection
'
'    For Each Rng In SelRange
'        d = Mid(Rng.Value, 4, 2)
'        m = Left(Rng.Value, 2)
'        y = Right(Rng.Value, 4)
'        Rng.Value = Format(DateSerial(y, m, d), "yyyy-mm-dd")
'    Next Rng
'End Sub

Sub Insert_Trades_in_MySQL()
    
    Application.DecimalSeparator = "."
    Application.ThousandsSeparator = ","
    Application.UseSystemSeparators = True
        
    Dim Server_Name As String
    Dim Database_Name As String
    Dim User_ID As String
    Dim Password As String
    
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    ipddbb = "127.0.0.1"
    nameddbb = ""
    userddbb = ""
    passddbb = ""
    
    If Not IsEmpty([server]) Then
        ipddbb = [server]
    End If
    If Not IsEmpty([ddbb]) Then
        nameddbb = [ddbb]
    End If
    If Not IsEmpty([user_mysql]) Then
        userddbb = [user_mysql]
    End If
    If Not IsEmpty([pass_mysql]) Then
        passddbb = [pass_mysql]
    End If
    
    Set ActSheet = ActiveSheet
    Set SelRange = Selection
    
    Set conn = New ADODB.Connection
        
    conn.Open "DRIVER={MySQL ODBC 5.2 Unicode Driver};" & _
                "SERVER=" & ipddbb & ";" & _
                "PORT=3306;" & _
                "DATABASE=" & nameddbb & ";" & _
                "USER=" & userddbb & ";" & _
                "PASSWORD=" & passddbb & ";" & _
                "OPTION=3;"""
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    
    the_table = "ddbbname.trades"
    
    If SelRange.Rows.Count >= 2 Then

        resp = MsgBox("Insert into MySQL:" & vbCrLf & _
                      "<Yes> Trade File to PNC" & vbCrLf & _
                      "<No> Additional Trade File to BNY", _
                      vbYesNoCancel, _
                      "Trades to MySQL")
        
        If resp <> 6 And resp <> 7 Then
            Exit Sub
        End If
        
        For i = 2 To SelRange.Rows.Count
            
            ticker = SelRange(i, 1)
            idtrade = SelRange(i, 4)
            account = SelRange(i, 5)
            
            If resp = 6 Then
                
                security_id = Replace(SelRange(i, 6), "'", "\'")
                security_desc = Replace(SelRange(i, 7), "'", "\'")
                trade_date = Format(chgDateFormat(SelRange(i, 8)), "yyyy-mm-dd")
                settle_date = Format(chgDateFormat(SelRange(i, 9)), "yyyy-mm-dd")
                quantity = Replace(Format(CDbl(SelRange(i, 10)), "#0.0000"), ",", ".")
                broker = Replace(SelRange(i, 11), "'", "\'")
                buy_sell = Replace(SelRange(i, 12), "'", "\'")
                longshort = Replace(SelRange(i, 13), "'", "\'")
                net_amount = Replace(Format(CDbl(SelRange(i, 14)), "#0.0000"), ",", ".")
                px = Replace(Format(CDbl(SelRange(i, 15)), "#0.0000"), ",", ".")
                ccy = Replace(SelRange(i, 16), "'", "\'")
                sts = Replace(SelRange(i, 17), "'", "\'")
                record_type = Replace(SelRange(i, 18), "'", "\'")
                commission = Replace(Format(CDbl(SelRange(i, 19)), "#0.0000"), ",", ".")
                cusip = Replace(SelRange(i, 20), "'", "\'")
                sec_fee = Replace(Format(CDbl(SelRange(i, 21)), "#0.0000"), ",", ".")
                other_charges = Replace(Format(CDbl(SelRange(i, 22)), "#0.0000"), ",", ".")
                
                security_type = ""
                buy_currency = ""
                sell_currency = ""
                buy_ccy_amount = 0
                sell_ccy_amount = 0
            
            ElseIf resp = 7 Then
            
                security_type = Replace(SelRange(i, 6), "'", "\'")
                security_id = Replace(SelRange(i, 7), "'", "\'")
                security_desc = Replace(SelRange(i, 8), "'", "\'")
                trade_date = Format(chgDateFormat(SelRange(i, 9)), "yyyy-mm-dd")
                settle_date = Format(chgDateFormat(SelRange(i, 10)), "yyyy-mm-dd")
                quantity = Replace(Format(CDbl(SelRange(i, 11)), "#0.0000"), ",", ".")
                broker = Replace(SelRange(i, 12), "'", "\'")
                buy_sell = Replace(SelRange(i, 13), "'", "\'")
                longshort = Replace(SelRange(i, 14), "'", "\'")
                net_amount = Replace(Format(CDbl(SelRange(i, 15)), "#0.0000"), ",", ".")
                px = Replace(Format(CDbl(SelRange(i, 16)), "#0.0000"), ",", ".")
                ccy = Replace(SelRange(i, 17), "'", "\'")
                sts = Replace(SelRange(i, 18), "'", "\'")
                record_type = Replace(SelRange(i, 19), "'", "\'")
                commission = Replace(Format(CDbl(SelRange(i, 20)), "#0.0000"), ",", ".")
                cusip = Replace(SelRange(i, 21), "'", "\'")
                sec_fee = Replace(Format(CDbl(SelRange(i, 22)), "#0.000"), ",", ".")
                other_charges = Replace(Format(CDbl(SelRange(i, 23)), "#0.0000"), ",", ".")
                buy_currency = Replace(SelRange(i, 24), "'", "\'")
                sell_currency = Replace(Replace(SelRange(i, 25), "'", "\'"), ",", ".")
                buy_ccy_amount = Replace(Format(CDbl(SelRange(i, 26)), "#0.0000"), ",", ".")
                sell_ccy_amount = Replace(Format(CDbl(SelRange(i, 27)), "#0.0000"), ",", ".")
                
            End If
            
            strSQL = "INSERT INTO " & the_table & _
                        " (id, trade_date, settle_date,security_id, ticker, " & _
                        "security_type, security_desc, quantity, broker, buy_sell, " & _
                        "position, net_amount, price, currency, account, sts, record_type, " & _
                        "commission, cusip, sec_fee, other_charges, buy_currency, " & _
                        "sell_currency, buy_ccy_amount, sell_ccy_amount) "
            strSQL = strSQL & " VALUES ('" & _
                        idtrade & "','" & trade_date & "','" & settle_date & "','" & security_id & "','" & _
                        ticker & "','" & security_type & "','" & security_desc & "'," & quantity & ",'" & _
                        broker & "','" & buy_sell & "','" & longshort & "'," & _
                        net_amount & "," & px & ",'" & ccy & "','" & account & "','" & sts & "','" & _
                        record_type & "'," & commission & ",'" & cusip & "'," & sec_fee & "," & other_charges & ",'" & _
                        buy_currency & "','" & sell_currency & "'," & buy_ccy_amount & "," & sell_ccy_amount & ")"
                        
            strSQL = strSQL & " ON DUPLICATE KEY UPDATE " & _
                        "trade_date = '" & trade_date & "', " & "settle_date = '" & settle_date & "', " & _
                        "security_id = '" & security_id & "', " & "ticker = '" & ticker & "', " & _
                        "security_type = '" & security_type & "', " & "security_desc = '" & security_desc & "', " & _
                        "quantity = " & quantity & ", " & "broker = '" & broker & "', " & _
                        "buy_sell = '" & buy_sell & "', " & "position = '" & longshort & "', " & _
                        "net_amount = " & net_amount & ", " & "price = " & px & ", " & _
                        "currency = '" & ccy & "', " & "account = '" & account & "', " & _
                        "sts = '" & sts & "', " & "record_type = '" & record_type & "', " & _
                        "commission = " & commission & ", " & "cusip = '" & cusip & "', " & _
                        "sec_fee = " & sec_fee & ", " & "other_charges = " & other_charges & ", " & _
                        "buy_currency = '" & buy_currency & "', " & "sell_currency = '" & sell_currency & "', " & _
                        "buy_ccy_amount = " & buy_ccy_amount & ", " & "sell_ccy_amount = " & sell_ccy_amount & ";"
            
            cmd.CommandText = strSQL
            cmd.Execute strSQL
            
        Next i
        
        MsgBox ("A total of " & CStr(SelRange.Rows.Count - 1) & " new records successfully inserted into MySQL.")
        
    End If
    
End Sub

Function chgDateFormat(chgdate As Variant)
        d = Mid(chgdate, 4, 2)
        m = Left(chgdate, 2)
        y = Right(chgdate, 4)
        chgDateFormat = Format(DateSerial(y, m, d), "yyyy-mm-dd")
End Function


Sub checkInBloombergTbl()

    Dim conn As ADODB.Connection
    Dim Server_Name As String
    Dim Database_Name As String
    Dim User_ID As String
    Dim Password As String
    Dim SQLStr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
    
    Dim ActSheet As Worksheet
    Dim SelRange As Range
    
    Server_Name = "localhost" ' Enter your server name here
    Database_Name = "" ' Enter your database name here
    User_ID = "" ' enter your user ID here
    Password = "" ' Enter your password here
    
    Set ActSheet = ActiveSheet
    Set SelRange = Selection
    
    Set conn = New ADODB.Connection
    conn.Open "Driver={MySQL ODBC 5.2w Driver};Server=" & Server_Name & _
            ";Database=" & Database_Name & _
            ";Uid=" & User_ID & _
            ";Pwd=" & Password & ";"
    
    the_table = "ddbbname.bloomberg"
    
    For i = 1 To SelRange.Rows.Count
        
        notInTbl = 0
        
        strSQL = "SELECT COUNT(*) FROM " & the_table & " WHERE ticker = '"
        strSQL = strSQL & SelRange(i, 1) & "';"
        
        rs.Open strSQL, conn, adOpenStatic
        
        notInTbl = rs.GetRows()
        
        If notInTbl(0, 0) = 0 Then
            MsgBox (SelRange(i, 1) & " is not in " & the_table)
        End If
        
        rs.Close
        
    Next i
    
    Set rs = Nothing
    
    conn.Close
    Set conn = Nothing

    MsgBox ("All checked!")

End Sub


Sub DuplicateValue()

    Application.ScreenUpdating = False
    
    Dim ActSheet As Worksheet
    Dim SelRange As Range
    Dim i As Integer, j As Integer
    
    Set ActSheet = ActiveSheet
    Set SelRange = Selection
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
    r1 = SelRange.Row
    r2 = (SelRange.Row + SelRange.Rows.Count - 1)
    c1 = SelRange.Column
    c2 = (SelRange.Column + SelRange.Columns.Count - 1)
    
    'checking on first worksheet
    r = r1
    c = c1
    While r <= r2
        While c <= c2
            If (WorksheetFunction.CountIf(Sheets(ActiveSheet.Name).Range(Cells(r, c1), Cells(r, c)), Cells(r, c).Value) > 1) Or _
                IsError(Cells(r, c).Value) Then
                Range(Cells(r, c), Cells(r, c)).Select
                Selection.Delete Shift:=xlToLeft
            Else
                c = c + 1
            End If
        Wend
        r = r + 1
        c = c1
    Wend
    
    Application.ScreenUpdating = True
    
End Sub

