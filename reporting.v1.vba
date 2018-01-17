Sub cleardata()

    Dim sht As Worksheet
    Set sht = ActiveSheet
    lastrow = getLastRow
    
    If lastrow > 10 Then
        resp = MsgBox("Delete " & CStr(lastrow - 10) & " rows?" & vbCrLf & _
                          "<Yes> Delete" & vbCrLf & _
                          "<No> Abort" & vbCrLf, _
                          vbYesNo, _
                          "Delete confirmation")
            
        If resp = 7 Then
            Exit Sub
        End If
        
        If sht.Name = "Assets" Then
            With sht
                .Range("A11:X" & lastrow).ClearContents
                .Range("A11:X" & lastrow).NumberFormat = "General"
            End With
        ElseIf sht.Name = "Hist." Then
            With sht
                .Range("A11:E" & lastrow).ClearContents
            End With
        Else
            With sht
                .Range("A11:P" & lastrow).ClearContents
            End With
        End If
    End If
End Sub

Public Function blotter_formatCells()

    Dim sht As Worksheet
    Dim col2TextFormat() As Variant
    Dim col2NumFormat() As Variant
    Dim col2DblFormat() As Variant
    Dim col2TimeFormat() As Variant
    
    Set sht = ActiveSheet
    'Format Columns
    col2TextFormat = Array(1, 2, 4, 5, 7, 8, 10, 12, 14, 15) ' Columns in excel with text format (i.e.: ID, Date, Type, Position, ...)
    col2IntgFormat = Array(6, 11, 16)                 ' Columns in excel with num format (i.e.: Limit, Price)
    col2DblFormat = Array(9, 13)                ' Columns in excel with num format (i.e.: Amount, Filled, Remainder)
    col2TimeFormat = Array(3)                ' Columns in excel with num format (i.e.: Time)
    
    lastrow = getLastRow
    
    For col = 1 To 15
        If findinarray(col2TextFormat, col) Then
            sht.Range(sht.Cells(11, col), sht.Cells(lastrow, col)).NumberFormat = "@"
            If col = 2 Or col = 4 Or col = 5 Or col = 7 Or col = 12 Or col = 14 Then
                sht.Range(sht.Cells(11, col), sht.Cells(lastrow, col)).HorizontalAlignment = xlCenter
            Else
                sht.Range(sht.Cells(11, col), sht.Cells(lastrow, col)).HorizontalAlignment = xlLeft
            End If
        ElseIf findinarray(col2IntgFormat, col) Then
            sht.Range(sht.Cells(11, col), sht.Cells(lastrow, col)).NumberFormat = "#,##0;[Red]-#,##0"
            sht.Range(sht.Cells(11, col), sht.Cells(lastrow, col)).HorizontalAlignment = xlRight
        ElseIf findinarray(col2DblFormat, col) Then
            sht.Range(sht.Cells(11, col), sht.Cells(lastrow, col)).NumberFormat = "#,##0.0000;[Red]-#,##0.0000"
            sht.Range(sht.Cells(11, col), sht.Cells(lastrow, col)).HorizontalAlignment = xlRight
        ElseIf findinarray(col2TimeFormat, col) Then
            sht.Range(sht.Cells(11, col), sht.Cells(lastrow, col)).NumberFormat = "hh:mm:ss"
            sht.Range(sht.Cells(11, col), sht.Cells(lastrow, col)).HorizontalAlignment = xlCenter
        End If
    Next col
End Function

Public Function findinarray(arr, searchterm) As Boolean
    
    For i = 0 To UBound(arr, 1)
        If arr(i) = searchterm Then GoTo bypass
    Next i
    findinarray = False
    Exit Function
    
bypass:
    findinarray = True
    
End Function

Public Function randomStr() As String

    Dim maxlen As Integer
    Dim upperbound As Integer
    Dim lowerbound As Integer
    
    Application.Volatile
    
    upperbound = 65
    lowerbound = 90
    maxlen = 6
    randomStr = ""
    For i = 1 To maxlen
        Randomize
        randomStr = randomStr & Chr(Int((upperbound - lowerbound + 1) * Rnd + lowerbound))
    Next i
    randomStr = Format(Date, "yyyymmdd") & "-" & randomStr

End Function

Public Function isvaliddate(date2validate As String) As Boolean
    'date2validate with format "yyyy-mm-dd"
    yy = Left(date2validate, 4)
    mm = Mid(date2validate, 5, 2)
    dd = Right(date2validate, 2)
    
    isvaliddate = False
    If Len(date2validate) = 10 Then
        If IsNumeric(yy) And IsNumeric(mm) And IsNumeric(dd) Then
            If mm > 0 Or mm <= 12 And dd > 0 Or dd <= 31 Then
                isvaliddate = True
            End If
        End If
    End If
End Function

Public Function getLastRow()

    Dim sht As Worksheet
    Set sht = ActiveSheet
    lastrow = sht.Cells(sht.rows.Count, "A").End(xlUp).Row
    lastRowF = sht.Cells(sht.rows.Count, "F").End(xlUp).Row
    If lastRowF > lastrow Then
        lastrow = lastRowF
    End If
    If lastrow <= 10 Then
        lastrow = 10
    End If
    getLastRow = lastrow
End Function

Public Function blotter_validateFields()

    Dim sht As Worksheet
    Dim msg As String
    
    Set sht = ActiveSheet
    lastrow = getLastRow

    msg = ""
    For j = 11 To lastrow
        
        If sht.Cells(j, 1) = "" Then
            msg = "Check ID in row: " & CStr(j) & ". ID cannot be blank"
            Exit For
        End If
        If sht.Cells(j, 2) = "" Then
            msg = "Check date in row: " & CStr(j) & ". Date cannot be blank"
            Exit For
        ElseIf Not isvaliddate(sht.Cells(j, 2)) Then
            msg = "Check date in row: " & CStr(j) & ". It is not a valid date"
            Exit For
        End If
        t = Format(sht.Cells(j, 3), "hh:mm")
        If sht.Cells(j, 3) = "" Then
            msg = "Check time in row: " & CStr(j) & ". Time cannot be blank"
            Exit For
        ElseIf Len(t) <> 5 Or Mid(t, 3, 1) <> ":" Or _
            Left(t, 2) > 23 Or Left(t, 2) < 0 Or _
            Right(t, 2) > 59 Or Right(t, 2) < 0 Then
            msg = "Check time in row: " & CStr(j) & ". It is not a valid time"
            Exit For
        End If
        If Not sht.Cells(j, 4) = "BUY" And Not sht.Cells(j, 4) = "SELL" Then
            msg = "Check type in row: " & CStr(j) & ". It is only valid either BUY or SELL"
            Exit For
        End If
        If Not sht.Cells(j, 5) = "LONG" And Not sht.Cells(j, 5) = "SHORT" Then
            msg = "Check type in row: " & CStr(j) & ". It is only valid either LONG or SHORT"
            Exit For
        End If
        If sht.Cells(j, 6) = "" Then
            msg = "Check amount in row: " & CStr(j) & ". The amount cannot be blank"
            Exit For
        ElseIf Not IsNumeric(sht.Cells(j, 6)) Then
            msg = "Check amount in row: " & CStr(j) & ". The amount is not a number"
            Exit For
        ElseIf sht.Cells(j, 6) < 0 And sht.Cells(j, 4) = "BUY" Or _
                sht.Cells(j, 6) > 0 And sht.Cells(j, 4) = "SELL" Then
            msg = "Check amount in row: " & CStr(j) & ". BUY/SELL order but NEG/POS amount number"
            Exit For
        End If
        If sht.Cells(j, 7) = "" Then
            msg = "Check ticker in row: " & CStr(j) & ". Ticker cannot be blank"
            Exit For
        End If
        res = selectfromddbb("ticker", "assets", "where ticker = '" & sht.Cells(j, 7) & "'")
        If sht.Cells(j, 7) = "" Then
            msg = "Check name in row: " & CStr(j) & ". Name cannot be blank"
            Exit For
        ElseIf res <> sht.Cells(j, 7) Then
            msg = "Check name in row: " & CStr(j) & ". The ticker is not in the database. Please, update it first before continue."
            Exit For
        End If
        If Not IsNumeric(sht.Cells(j, 9)) And sht.Cells(j, 9) <> "" Then
            msg = "Check limit in row: " & CStr(j) & ". Limit must be a number or leave it blank"
            Exit For
        End If
        If Not IsNumeric(sht.Cells(j, 11)) Then
            msg = "Check filled in row: " & CStr(j) & ". Filled must be a number"
            Exit For
        ElseIf sht.Cells(j, 11) < 0 And sht.Cells(j, 4) = "BUY" Or _
                sht.Cells(j, 11) > 0 And sht.Cells(j, 4) = "SELL" Then
            msg = "Check filled in row: " & CStr(j) & ". BUY/SELL order but NEG/POS filled number"
            Exit For
        End If
        res = selectfromddbb("ccy", "assets", "where ticker = '" & sht.Cells(j, 7) & "'")
        If res <> sht.Cells(j, 12) Then
            msg = "Check currency in row: " & CStr(j) & ". Currency cannot be blank"
            Exit For
        ElseIf res = "#--" Then
            msg = "Check currency in row: " & CStr(j) & ". No ccy has been set for this ticker in the database. Please, update it first before continue."
            Exit For
        End If
        If Not IsNumeric(sht.Cells(j, 13)) Then
            msg = "Check price in row: " & CStr(j) & ". Price must be a number"
            Exit For
        ElseIf sht.Cells(j, 13) < 0 Then
            msg = "Check Price in row: " & CStr(j) & ". Price cannot be negtive"
            Exit For
        End If
        If sht.Cells(j, 14) = "" Then
            msg = "Check broker in row: " & CStr(j) & ". Please, specify the broker."
            Exit For
        End If
    Next j
    
    blotter_validateFields = msg

End Function

Public Function hist_validateFields()
    Dim sht As Worksheet
    Dim msg As String
    
    Set sht = ActiveSheet
    lastrow = getLastRow

    msg = ""
    For j = 11 To lastrow
        
        If sht.Cells(j, 1) = "" Then
            msg = "Check Ticker in row: " & CStr(j) & ". Ticker cannot be blank"
            Exit For
        End If
        If sht.Cells(j, 2) = "" Then
            msg = "Check Type in row: " & CStr(j) & ". Type cannot be blank"
            Exit For
        End If
        If sht.Cells(j, 3) = "" Then
            msg = "Check Date in row: " & CStr(j) & ". Date cannot be blank"
            Exit For
        ElseIf Not isvaliddate(sht.Cells(j, 3)) Then
            msg = "Check Date in row: " & CStr(j) & ". Date is not yyyy-mm-dd"
            Exit For
        End If
        If sht.Cells(j, 4) = "" Then
            msg = "Check Price in row: " & CStr(j) & ". Price cannot be blank"
            Exit For
        ElseIf Not IsNumeric(sht.Cells(j, 4)) Then
            msg = "Check Price in row: " & CStr(j) & ". Price is not a number"
            Exit For
        End If
        If sht.Cells(j, 5) = "" Then
            msg = "Check Volume in row: " & CStr(j) & ". Volume cannot be blank"
            Exit For
        ElseIf Not IsNumeric(sht.Cells(j, 5)) Then
            msg = "Check Volume in row: " & CStr(j) & ". Volume is not a number"
            Exit For
        End If
    Next j
    hist_validateFields = msg
End Function

Sub blotter_save_in_MySQL()

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
    Dim rng As Range
    Dim c As Range
    
    Set sht = ActiveSheet
    
    ipddbb = ""
    portddbb = 3306
    nameddbb = ""
    userddbb = ""
    passddbb = ""
    
    If Not IsEmpty([mysql_server]) Then
        ipddbb = [mysql_server]
    End If
    If Not IsEmpty([mysql_ddbb]) Then
        nameddbb = [mysql_ddbb]
    End If
    If Not IsEmpty([mysql_port]) Then
        portddbb = [mysql_port]
    End If
    If Not IsEmpty([mysql_user]) Then
        userddbb = [mysql_user]
    End If
    If Not IsEmpty([mysql_pass]) Then
        passddbb = [mysql_pass]
    End If
    
    lastrow = getLastRow
    
    Set rng = sht.Range("A10:O" & lastrow)
    For Each c In rng
        c.Value = UCase(Trim(c.Value))
    Next c

    msg = blotter_validateFields
    If msg <> "" Then
        MsgBox (msg)
        Exit Sub
    End If
    
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
    
    tblddbb = "fundtrades"
    If sht.Name = "BLOTTER-SWAP" Then
        tblddbb = "swaptrades"
    End If
    
    If lastrow > 10 Then
        counter = 0
        resp = MsgBox("Insert into MySQL:" & vbCrLf & _
                      "<Yes> Save" & vbCrLf & _
                      "<No> Abort" & vbCrLf, _
                      vbYesNo, _
                      "Trades to MySQL")
        
        If resp = 7 Then
            Exit Sub
        End If
        
        For i = 11 To lastrow
            
            trd_id = sht.Cells(i, 1)                                                        ' ID
            trd_date = Replace(sht.Cells(i, 2), "'", "\'")                                  ' Date
            trd_time = Format(sht.Cells(i, 3), "hh:mm:ss")                                  ' Time
            trd_type = Replace(sht.Cells(i, 4), "'", "\'")                                  ' Type (Buy/Sell)
            trd_position = Replace(sht.Cells(i, 5), "'", "\'")                              ' Type (Long/Short)
            trd_amount = Replace(Format(CDbl(sht.Cells(i, 6)), "#0.0000"), ",", ".")        ' Amount
            trd_ticker = Replace(sht.Cells(i, 7), "'", "\'")                                ' Ticker
            trd_short_name = Replace(sht.Cells(i, 8), "'", "\'")                            ' Name
            trd_limt = Replace(Format(CDbl(sht.Cells(i, 9)), "#0.0000"), ",", ".")          ' Limit
            trd_filled = Replace(Format(CDbl(sht.Cells(i, 11)), "#0.0000"), ",", ".")       ' Filled
            trd_ccy = Replace(sht.Cells(i, 12), "'", "\'")                                  ' Currency
            trd_pxgross = Replace(Format(CDbl(sht.Cells(i, 13)), "#0.0000"), ",", ".")      ' PX (Gross)
            trd_broker = Replace(sht.Cells(i, 14), "'", "\'")                               ' Broker
            trd_instr = Replace(sht.Cells(i, 15), "'", "\'")                                ' Instruction
            trd_remaining = Replace(Format(CDbl(sht.Cells(i, 16)), "#0.0000"), ",", ".")    ' Remaining
            
            strSQL = "INSERT INTO " & tblddbb & _
                        " (trd_id, trd_date, trd_time, trd_type, trd_position, trd_amount, trd_ticker, " & _
                        "trd_short_name, trd_limit, trd_filled, " & _
                        "trd_ccy, trd_pricegross, trd_broker, trd_instructions) "
            strSQL = strSQL & " VALUES ('" & _
                        trd_id & "','" & trd_date & "','" & trd_time & "','" & _
                        trd_type & "','" & trd_position & "'," & trd_amount & ",'" & trd_ticker & "','" & _
                        trd_short_name & "'," & trd_limt & "," & _
                        trd_filled & ",'" & trd_ccy & "'," & trd_pxgross & ",'" & _
                        trd_broker & "','" & trd_instr & "')"
                        
            strSQL = strSQL & " ON DUPLICATE KEY UPDATE " & _
                        "trd_id = '" & trd_id & "', " & "trd_date = '" & trd_date & "', " & _
                        "trd_time = '" & trd_time & "', " & "trd_type = '" & trd_type & "', " & _
                        "trd_position = '" & trd_position & "', " & _
                        "trd_amount = " & trd_amount & ", " & "trd_ticker = '" & trd_ticker & "', " & _
                        "trd_short_name = '" & trd_short_name & "', " & "trd_limit = " & trd_limt & ", " & _
                        "trd_filled = " & trd_filled & ", " & _
                        "trd_ccy = '" & trd_ccy & "', " & "trd_pricegross = " & trd_pxgross & ", " & _
                        "trd_broker = '" & trd_broker & "', " & "trd_instructions = '" & trd_instr & "';"
            
            cmd.CommandText = strSQL
            Debug.Print (strSQL)
            
            cmd.Execute strSQL
            counter = counter + 1
        Next i
        
        MsgBox ("A total of " & CStr(counter) & " new records successfully inserted into MySQL.")
        
    End If
    
    conn.Close
    Set conn = Nothing

End Sub

Sub blotter_search()

    Dim conn As ADODB.Connection
    Dim SQLStr As String
    Dim rs As ADODB.Recordset
    
    Dim ipddbb As Variant
    Dim portddbb As Integer
    Dim nameddbb As Variant
    Dim userddbb As Variant
    Dim passddbb As Variant
    Dim fld As String
    Dim tbl As String
    
    Dim sht As Worksheet
    
    Application.ScreenUpdating = False
    
    Set sht = ActiveSheet
    lastrow = getLastRow + 1
      
    ipddbb = ""
    portddbb = 3306
    nameddbb = ""
    userddbb = ""
    passddbb = ""
    
    If Not IsEmpty([mysql_server]) Then
        ipddbb = [mysql_server]
    End If
    If Not IsEmpty([mysql_port]) Then
        portddbb = [mysql_port]
    End If
    If Not IsEmpty([mysql_ddbb]) Then
        nameddbb = [mysql_ddbb]
    End If
    If Not IsEmpty([mysql_user]) Then
        userddbb = [mysql_user]
    End If
    If Not IsEmpty([mysql_pass]) Then
        passddbb = [mysql_pass]
    End If
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    
    fldddbb = "trd_id, DATE_FORMAT(trd_date,'%Y-%m-%d'), DATE_FORMAT(trd_time,'%k:%i:%s'), " & _
          "trd_type, trd_position, trd_amount, trd_ticker, trd_short_name, " & _
          "trd_limit, trd_filled, trd_ccy, trd_pricegross, trd_broker, " & _
          "trd_instructions"
    
    tblddbb = "fundtrades"
    If sht.Name = "BLOTTER-SWAP" Then
        tblddbb = "swaptrades"
    End If
    
    strSQL = "SELECT " & fldddbb
    strSQL = strSQL & " FROM " & tblddbb & " "
    
    whereStr = False
    If sht.Name = "BLOTTER-SWAP" Then
    
        If Not IsEmpty([rtn_swp_id]) Then
            whereStr = True
            strSQL = strSQL & "WHERE trd_id='" & [rtn_swp_id] & "' "
        End If
        If Not IsEmpty([rtn_swp_from]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_date>='" & Format([rtn_swp_from], "yyyy-mm-dd") & "' "
            Else
                strSQL = strSQL & "WHERE trd_date>='" & Format([rtn_swp_from], "yyyy-mm-dd") & "' "
            End If
            whereStr = True
        End If
        If Not IsEmpty([rtn_swp_to]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_date<='" & Format([rtn_swp_to], "yyyy-mm-dd") & "' "
            Else
                strSQL = strSQL & "WHERE trd_date<='" & Format([rtn_swp_to], "yyyy-mm-dd") & "' "
            End If
            whereStr = True
        End If
        If Not IsEmpty([rtn_swp_ticker]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_ticker='" & [rtn_swp_ticker] & "' "
            Else
                strSQL = strSQL & "WHERE trd_ticker='" & [rtn_swp_ticker] & "' "
            End If
            whereStr = True
        End If
        If Not IsEmpty([rtn_swp_type]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_type = '" & [rtn_swp_type] & "' "
            Else
                strSQL = strSQL & "WHERE trd_type = '" & [rtn_swp_type] & "' "
            End If
            whereStr = True
        End If
        If Not IsEmpty([rtn_swp_position]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_position = '" & [rtn_swp_position] & "' "
            Else
                strSQL = strSQL & "WHERE trd_position = '" & [rtn_swp_position] & "' "
            End If
            whereStr = True
        End If
        If Not IsEmpty([rtn_swp_ccy]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_ccy='" & [rtn_swp_ccy] & "' "
            Else
                strSQL = strSQL & "WHERE trd_ccy='" & [rtn_swp_ccy] & "' "
            End If
            whereStr = True
        End If
        If Not IsEmpty([rtn_swp_bkr]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_broker='" & [rtn_swp_bkr] & "' "
            Else
                strSQL = strSQL & "WHERE trd_broker='" & [rtn_swp_bkr] & "' "
            End If
        End If
    Else
        If Not IsEmpty([rtn_fnd_id]) Then
            whereStr = True
            strSQL = strSQL & "WHERE trd_id='" & [rtn_fnd_id] & "' "
        End If
        If Not IsEmpty([rtn_fnd_from]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_date>='" & Format([rtn_fnd_from], "yyyy-mm-dd") & "' "
            Else
                strSQL = strSQL & "WHERE trd_date>='" & Format([rtn_fnd_from], "yyyy-mm-dd") & "' "
            End If
            whereStr = True
        End If
        If Not IsEmpty([rtn_fnd_to]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_date<='" & Format([rtn_fnd_to], "yyyy-mm-dd") & "' "
            Else
                strSQL = strSQL & "WHERE trd_date<='" & Format([rtn_fnd_to], "yyyy-mm-dd") & "' "
            End If
            whereStr = True
        End If
        If Not IsEmpty([rtn_fnd_ticker]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_ticker='" & [rtn_fnd_ticker] & "' "
            Else
                strSQL = strSQL & "WHERE trd_ticker='" & [rtn_fnd_ticker] & "' "
            End If
            whereStr = True
        End If
        If Not IsEmpty([rtn_fnd_type]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_type = '" & [rtn_fnd_type] & "' "
            Else
                strSQL = strSQL & "WHERE trd_type = '" & [rtn_fnd_type] & "' "
            End If
            whereStr = True
        End If
        If Not IsEmpty([rtn_fnd_position]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_position = '" & [rtn_fnd_position] & "' "
            Else
                strSQL = strSQL & "WHERE trd_position = '" & [rtn_fnd_position] & "' "
            End If
            whereStr = True
        End If
        If Not IsEmpty([rtn_fnd_ccy]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_ccy='" & [rtn_fnd_ccy] & "' "
            Else
                strSQL = strSQL & "WHERE trd_ccy='" & [rtn_fnd_ccy] & "' "
            End If
            whereStr = True
        End If
        If Not IsEmpty([rtn_fnd_bkr]) Then
            If whereStr = True Then
                strSQL = strSQL & "AND trd_broker='" & [rtn_fnd_bkr] & "' "
            Else
                strSQL = strSQL & "WHERE trd_broker='" & [rtn_fnd_bkr] & "' "
            End If
        End If
    End If
    
    strSQL = strSQL & "ORDER BY trd_date, trd_time, trd_ticker ASC"
    
    conn.Open "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
                "SERVER=" & ipddbb & ";" & _
                "PORT=" & CStr(portddbb) & ";" & _
                "DATABASE=" & nameddbb & ";" & _
                "USER=" & userddbb & ";" & _
                "PASSWORD=" & passddbb & ";" & _
                "OPTION=3;"""
    
    rs.Open strSQL, conn, adOpenStatic
    
    '----------------------------------------------
    Dim myArray()
    
    MySQLtoExcel = "#--"
    
    Do While Not rs.EOF
    
        myArray = rs.GetRows()
    
        m = UBound(myArray, 2)
        n = UBound(myArray, 1)
                
        i = 0
        j = 0
        k = 0
        Do While i <= m ' Using For loop data are displayed
            Do While j <= n
                thisdata = myArray(j, i)
                If thisdata = "#N/A" Then
                    Range("A" & lastrow).Offset(i, k).Value = "--"
                Else
                    Range("A" & lastrow).Offset(i, k).Value = thisdata
                End If
                j = j + 1
                k = k + 1
                If k = 9 Then  ' No status
                    k = k + 1
                End If
                thisdata = ""
            Loop
            i = i + 1
            j = 0
            k = 0
        Loop
        
    Loop
        '----------------------------------------------

    rs.Close
    Set rs = Nothing
    
    conn.Close
    Set conn = Nothing
    
    blotter_formatCells
    Application.ScreenUpdating = True
    
End Sub

Sub blotter_print2pdf()
    
    Dim sht As Worksheet
    Dim strPath As String
    Dim myFile As Variant
    Dim strFile As String
    On Error GoTo errHandler

    Application.ScreenUpdating = False
    
    Set sht = ActiveSheet
    lastrow = getLastRow

    sht.Columns("B:B").ColumnWidth = 17.85   ' ID
    sht.Columns("B:B").ColumnWidth = 13.73   ' DATE
    sht.Columns("C:C").ColumnWidth = 16.5    ' TIME
    sht.Columns("D:D").ColumnWidth = 13.73   ' TYPE
    sht.Columns("E:E").ColumnWidth = 13.73   ' POSITION
    sht.Columns("F:F").ColumnWidth = 16.5    ' AMOUNT
    sht.Columns("G:G").ColumnWidth = 17.85   ' TICKER
    sht.Columns("H:H").ColumnWidth = 22      ' NAME
    sht.Columns("I:I").ColumnWidth = 16.5    ' LIMIT
    sht.Columns("J:J").ColumnWidth = 30.2    ' STATUS
    sht.Columns("K:K").ColumnWidth = 16.5    ' FILLED
    sht.Columns("L:L").ColumnWidth = 12      ' CCY
    sht.Columns("M:M").ColumnWidth = 16.5    ' PRICE
    sht.Columns("N:N").ColumnWidth = 16.5    ' BROKER
    sht.Columns("O:O").ColumnWidth = 50      ' INSTRUCTIONS
    sht.Columns("P:P").ColumnWidth = 16.5    ' REMAINING
    
    With sht.PageSetup
        .PrintArea = "$B$10:$O$" & lastrow
        .PrintTitleRows = "$B$10:$O$10"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.6)
        .BottomMargin = Application.InchesToPoints(0.6)
        .CenterHeader = "&B" & sht.Name & "&B  -  &D"
        .RightFooter = "Page &P of &N"
    End With
    
    'enter name and select folder for file
    ' start in current workbook folder
    strFile = Replace(Replace(sht.Name, " ", ""), ".", "_") _
                & "_" _
                & Format(Now(), "yyyymmdd\_hhmmss") _
                & ".pdf"
    strFile = ThisWorkbook.Path & "\" & strFile
    
    myFile = Application.GetSaveAsFilename _
        (InitialFileName:=strFile, _
            FileFilter:="PDF Files (*.pdf), *.pdf", _
            Title:="Select Folder and FileName to save")
    
    If myFile <> "False" Then
        sht.Columns("C").Hidden = True          ' TIME
        sht.Columns("J").Hidden = True          ' STATUS
        sht.Range("A10:P" & lastrow).Font.Size = 14
        sht.Columns("A:P").WrapText = True
        sht.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=myFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        sht.Columns("C").Hidden = False         ' TIME
        sht.Columns("J").Hidden = False         ' STATUS
        sht.Columns("A:P").WrapText = False
        sht.Range("A10:P" & lastrow).Font.Size = 11
        
        With sht.PageSetup
            .PrintArea = ""
            .PrintTitleRows = ""
        End With
        MsgBox "PDF file has been created."
    End If
    
    Application.ScreenUpdating = True
    
exitHandler:
        Exit Sub
errHandler:
        MsgBox "cannot create PDF file"
        Resume exitHandler
End Sub

Sub blotter_send_email()
   
    Dim intChoice As Integer
    Dim strPath As String
    Dim sht As Worksheet
    
    Set sht = ActiveSheet
    'only allow the user to select one file
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    'make the file dialog visible to the user
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    'determine what choice the user made
    If intChoice <> 0 Then
        'get the file path selected by the user
        pdfFileName = Application.FileDialog( _
            msoFileDialogOpen).SelectedItems(1)
         
        Set Mail_Object = CreateObject("Outlook.Application")
             With Mail_Object.CreateItem(o)
                 .Subject = sht.Name & " - " & Format(Date, "yyyy-mm-dd")
                 .To = "email@company.com"
                 .Body = "Attached today's blotter." & Chr(13) & _
                         "Have a nice day!" & Chr(13) & Chr(13) & _
                         "Best regards," & Chr(13) & _
                         "Admin"
                 .Attachments.Add pdfFileName
                 .Send
         End With
             MsgBox "E-mail successfully sent", 64
             Application.DisplayAlerts = False
             
         Set Mail_Object = Nothing
    End If
End Sub

Sub assets_search()
    Dim conn As ADODB.Connection
    Dim SQLStr As String
    Dim rs As ADODB.Recordset
    
    Dim ipddbb As Variant
    Dim portddbb As Integer
    Dim nameddbb As Variant
    Dim userddbb As Variant
    Dim passddbb As Variant
    Dim fld As String
    Dim tbl As String
    
    Dim sht As Worksheet
    
    Set sht = ActiveSheet
    
    lastrow = getLastRow + 1
    If lastrow < 11 Then
        lastrow = 11
    End If
      
    ipddbb = ""
    portddbb = 3306
    nameddbb = ""
    userddbb = ""
    passddbb = ""
    tblddbb = "assets"
    
    If Not IsEmpty([mysql_server]) Then
        ipddbb = [mysql_server]
    End If
    If Not IsEmpty([mysql_port]) Then
        portddbb = [mysql_port]
    End If
    If Not IsEmpty([mysql_ddbb]) Then
        nameddbb = [mysql_ddbb]
    End If
    If Not IsEmpty([mysql_user]) Then
        userddbb = [mysql_user]
    End If
    If Not IsEmpty([mysql_pass]) Then
        passddbb = [mysql_pass]
    End If
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    
    fld = "ticker, sectype, sedol, isin, " & _
          "short_name, ccy, secdesc, country, " & _
          "eq_fundccy, eq_sector, eq_industry, eq_subindustry, " & _
          "opt_maturity, opt_strike, opt_valuepoint, opt_underlying, " & _
          "fut_underlying, fut_valuepoint, fut_tickvalue, fut_ticksize, " & _
          "fut_contractsize, fut_firsttradedate, fut_lasttradedate, isvalid "
    
    strSQL = "SELECT " & fld
    strSQL = strSQL & " FROM " & tblddbb & " "
    
    whereStr = False
    
    If Not IsEmpty([asset_ticker]) Then
        whereStr = True
        strSQL = strSQL & "WHERE ticker='" & [asset_ticker] & "' "
    End If
    If Not IsEmpty([asset_sedol]) Then
        If whereStr = True Then
            strSQL = strSQL & "AND sedol='" & [asset_sedol] & "' "
        Else
            strSQL = strSQL & "WHERE sedol='" & [asset_sedol] & "' "
        End If
        whereStr = True
    End If
    If Not IsEmpty([asset_isin]) Then
        If whereStr = True Then
            strSQL = strSQL & "AND isin='" & [asset_isin] & "' "
        Else
            strSQL = strSQL & "WHERE isin='" & [asset_isin] & "' "
        End If
        whereStr = True
    End If
    If Not IsEmpty([asset_name]) Then
        If whereStr = True Then
            strSQL = strSQL & "AND short_name='" & [asset_name] & "' "
        Else
            strSQL = strSQL & "WHERE short_name='" & [asset_name] & "' "
        End If
        whereStr = True
    End If
    If Not IsEmpty([asset_ccy]) Then
        If whereStr = True Then
            strSQL = strSQL & "AND ccy='" & [asset_ccy] & "' "
        Else
            strSQL = strSQL & "WHERE ccy='" & [asset_ccy] & "' "
        End If
        whereStr = True
    End If
    If Not IsEmpty([asset_desc]) Then
        If whereStr = True Then
            strSQL = strSQL & "AND secdesc LIKE '%" & [asset_desc] & "%' "
        Else
            strSQL = strSQL & "WHERE secdesc LIKE '%" & [asset_desc] & "%' "
        End If
        whereStr = True
    End If
    If Not IsEmpty([asset_ctry]) Then
        If whereStr = True Then
            strSQL = strSQL & "AND country='" & [asset_ctry] & "' "
        Else
            strSQL = strSQL & "WHERE country='" & [asset_ctry] & "' "
        End If
        whereStr = True
    End If
    If Not IsEmpty([asset_sector]) Then
        If whereStr = True Then
            strSQL = strSQL & "AND eq_sector='" & [asset_sector] & "' "
        Else
            strSQL = strSQL & "WHERE eq_sector='" & [asset_sector] & "' "
        End If
        whereStr = True
    End If
    If Not IsEmpty([asset_isvalid]) Then
        If whereStr = True Then
            strSQL = strSQL & "AND isvalid='" & [asset_isvalid] & "' "
        Else
            strSQL = strSQL & "WHERE isvalid='" & [asset_isvalid] & "' "
        End If
        whereStr = True
    End If
    
    
    strSQL = strSQL & " ORDER BY ticker ASC;"
    
    conn.Open "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
                "SERVER=" & ipddbb & ";" & _
                "PORT=" & CStr(portddbb) & ";" & _
                "DATABASE=" & nameddbb & ";" & _
                "USER=" & userddbb & ";" & _
                "PASSWORD=" & passddbb & ";" & _
                "OPTION=3;"""
    
    rs.Open strSQL, conn, adOpenStatic
    
    '----------------------------------------------
    Dim myArray()
    
    MySQLtoExcel = "#--"
    
    Do While Not rs.EOF
    
        myArray = rs.GetRows()
    
        m = UBound(myArray, 2)
        n = UBound(myArray, 1)
                
        i = 0
        j = 0
        Do While i <= m ' Using For loop data are displayed
            Do While j <= n
                thisdata = myArray(j, i)
                If thisdata = "#N/A" Then
                    Range("A" & lastrow).Offset(i, j).Value = "#--"
                Else
                    Range("A" & lastrow).Offset(i, j).Value = thisdata
                End If
                j = j + 1
                thisdata = ""
            Loop
            i = i + 1
            j = 0
        Loop
        
    Loop
        '----------------------------------------------

    rs.Close
    Set rs = Nothing
    
    conn.Close
    Set conn = Nothing


End Sub

Sub assets_save()
    Application.DecimalSeparator = "."
    Application.ThousandsSeparator = ","
    Application.UseSystemSeparators = True
        
    Dim ipddbb As String
    Dim nameddbb As String
    Dim portddbb As Integer
    Dim userddbb As String
    Dim passddbb As String
    Dim msg As String
    Dim t As Date
    Dim i As Integer
    Dim j As Integer
    Dim strTicker As String
    
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Dim sht As Worksheet
    Dim rng As Range
    
    Set sht = ActiveSheet
      
    ipddbb = ""
    portddbb = 3306
    nameddbb = ""
    userddbb = ""
    passddbb = ""
    
    If Not IsEmpty([mysql_server]) Then
        ipddbb = [mysql_server]
    End If
    If Not IsEmpty([mysql_ddbb]) Then
        nameddbb = [mysql_ddbb]
    End If
    If Not IsEmpty([mysql_port]) Then
        portddbb = [mysql_port]
    End If
    If Not IsEmpty([mysql_user]) Then
        userddbb = [mysql_user]
    End If
    If Not IsEmpty([mysql_pass]) Then
        passddbb = [mysql_pass]
    End If
    
    strField = "PX_LAST"
    If Not IsEmpty([hist_field]) Then
        strField = [hist_field]
    End If
    
    lastrow = getLastRow
    
    If lastrow > 10 Then
    
        resp = MsgBox("Insert into MySQL:" & vbCrLf & _
                      "<Yes> Save" & vbCrLf & _
                      "<No> Abort" & vbCrLf, _
                      vbYesNo, _
                      "Assets to MySQL")
        
        If resp = 7 Then
            Exit Sub
        End If
    
        Set rng = sht.Range("A10:W" & lastrow)
        rng.NumberFormat = "@"

        For Each c In rng
            c.Value = UCase(Trim(c.Value))
        Next c
        tblddbb = "assets"
        counter = 0
        
        For j = 11 To lastrow
            If sht.Cells(j, 1) = "" Then
                MsgBox ("Ticker must not be blank. Please check row " & CStr(j))
                Exit Sub
            End If
            If sht.Cells(j, 2) = "" Then
                MsgBox ("Type must not be blank. Please check row " & CStr(j))
                Exit Sub
            End If
        Next j
        
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
        
        For i = 11 To lastrow
            
            ticker = Replace(sht.Cells(i, 1), "'", "\'")                                                         ' TICKER
            sectype = Replace(sht.Cells(i, 2), "'", "\'")                                                        ' TYPE
            If IsError(sht.Cells(i, 3)) Then
                sedol = UCase("#--")
            Else
                sedol = UCase(Replace(sht.Cells(i, 3), "'", "\'"))                                                 ' SEDOL
            End If
            If IsError(sht.Cells(i, 4)) Then
                isin = UCase("#--")
            Else
                isin = UCase(Replace(sht.Cells(i, 4), "'", "\'"))                                                           ' ISIN
            End If
            If IsError(sht.Cells(i, 5)) Then
                short_name = UCase("#--")
            Else
                short_name = UCase(Replace(sht.Cells(i, 5), "'", "\'"))                                                     ' SHORT NAME
            End If
            If IsError(sht.Cells(i, 6)) Then
                ccy = UCase("#--")
            Else
                ccy = UCase(Replace(sht.Cells(i, 6), "'", "\'"))                                                            ' CURRENCY
            End If
            If IsError(sht.Cells(i, 7)) Then
                secdesc = UCase("#--")
            Else
                secdesc = UCase(Replace(sht.Cells(i, 7), "'", "\'"))                                                        ' SEC DESCRIPTION
            End If
            If IsError(sht.Cells(i, 8)) Then
                country = UCase("#--")
            Else
                country = UCase(Replace(sht.Cells(i, 8), "'", "\'"))                                                        ' COUNTRY
            End If
            If IsError(sht.Cells(i, 9)) Then
                eq_fundccy = UCase("#--")
            Else
                eq_fundccy = UCase(Replace(sht.Cells(i, 9), "'", "\'"))                                                     ' EQ. FUND CCY
            End If
            If IsError(sht.Cells(i, 10)) Then
                eq_sector = UCase("#--")
            Else
                eq_sector = UCase(Replace(sht.Cells(i, 10), "'", "\'"))                                                     ' EQ. SECTOR
            End If
            If IsError(sht.Cells(i, 11)) Then
                eq_industry = UCase("#--")
            Else
                eq_industry = UCase(Replace(sht.Cells(i, 11), "'", "\'"))                                                   ' EQ. INDUSTRY
            End If
            If IsError(sht.Cells(i, 12)) Then
                eq_subindustry = "1970-01-01"
            Else
                eq_subindustry = UCase(Replace(sht.Cells(i, 12), "'", "\'"))                                                ' EQ. SUBINDUSTRY
            End If
            If IsError(sht.Cells(i, 13)) Then
                opt_maturity = UCase("1970-01-01")
            Else
                opt_maturity = UCase(Format(sht.Cells(i, 13), "yyyy-mm-dd"))                                            ' OPT. MATURITY
            End If
            If IsError(sht.Cells(i, 14)) Then
                opt_strike = -1
            Else
                If IsNumeric(sht.Cells(i, 14)) Then
                    opt_strike = UCase(Replace(Format(CDbl(sht.Cells(i, 14)), "#0.0000"), ",", "."))                        ' OPT. STRIKE
                Else
                    opt_strike = UCase(Replace(sht.Cells(i, 14), "'", "\'"))
                End If
            End If
            If IsError(sht.Cells(i, 15)) Then
                opt_valuepoint = -1
            Else
                If IsNumeric(sht.Cells(i, 15)) Then
                    opt_valuepoint = UCase(Replace(Format(CDbl(sht.Cells(i, 15)), "#0.0000"), ",", "."))                    ' OPT. VALUE 1 PT
                Else
                    opt_valuepoint = UCase(Replace(sht.Cells(i, 15), "'", "\'"))
                End If
            End If
            If IsError(sht.Cells(i, 16)) Then
                opt_underlying = UCase("#--")
            Else
                opt_underlying = UCase(Replace(sht.Cells(i, 16), "'", "\'"))                                                ' OPT. UNDERLYING
            End If
            If IsError(sht.Cells(i, 17)) Then
                fut_underlying = UCase("#--")
            Else
                fut_underlying = UCase(Replace(sht.Cells(i, 17), "'", "\'"))                                                ' FUT. UNDERLYING
            End If
            If IsError(sht.Cells(i, 18)) Then
                fut_valuepoint = -1
            Else
                If IsNumeric(sht.Cells(i, 18)) Then
                    fut_valuepoint = UCase(Replace(Format(CDbl(sht.Cells(i, 18)), "#0.0000"), ",", "."))                    ' FUT. VALUE 1 PT
                Else
                    fut_valuepoint = UCase(Replace(sht.Cells(i, 18), "'", "\'"))
                End If
            End If
            If IsError(sht.Cells(i, 19)) Then
                fut_tickvalue = -1
            Else
                If IsNumeric(sht.Cells(i, 19)) Then
                    fut_tickvalue = UCase(Replace(Format(CDbl(sht.Cells(i, 19)), "#0.0000"), ",", "."))                     ' FUT. TICK VALUE
                Else
                    fut_tickvalue = UCase(Replace(sht.Cells(i, 19), "'", "\'"))
                End If
            End If
            If IsError(sht.Cells(i, 20)) Then
                fut_ticksize = -1
            Else
                If IsNumeric(sht.Cells(i, 20)) Then
                    fut_ticksize = UCase(Replace(Format(CDbl(sht.Cells(i, 20)), "#0.0000"), ",", "."))                      ' FUT. TICK SIZE
                Else
                    fut_ticksize = UCase(Replace(sht.Cells(i, 20), "'", "\'"))
                End If
            End If
            If IsError(sht.Cells(i, 21)) Then
                fut_contractsize = -1
            Else
                If IsNumeric(sht.Cells(i, 21)) Then
                    fut_contractsize = UCase(Replace(Format(CDbl(sht.Cells(i, 21)), "#0.0000"), ",", "."))                 ' FUT. CONTRACT SIZE
                Else
                    fut_contractsize = UCase(Replace(sht.Cells(i, 21), "'", "\'"))
                End If
            End If
            If IsError(sht.Cells(i, 22)) Then
                fut_firsttradedate = "1970-01-01"
            Else
                fut_firsttradedate = UCase(Format(sht.Cells(i, 22), "yyyy-mm-dd"))                                         ' FUT. FIRST TRADE DATE
            End If
            If IsError(sht.Cells(i, 23)) Then
                fut_lasttradedate = "1970-01-01"
            Else
                fut_lasttradedate = UCase(Format(sht.Cells(i, 23), "yyyy-mm-dd"))                                          ' FUT. LAST TRADE DATE
            End If
            isvalidasset = 0
            If IsNumeric(sht.Cells(i, 24)) Then
                If sht.Cells(i, 24) = 0 Or sht.Cells(i, 24) = 1 Then
                    isvalidasset = UCase(Replace(Format(CDbl(sht.Cells(i, 24)), "#0.0000"), ",", "."))                     ' Is Valid
                End If
            End If
            
            
            If sedol = UCase("#N/A N/A") Or sedol = UCase("#N/A Invalid Security") Or sedol = UCase("#N/A Field Not Applicable") Or _
                sedol = UCase("#N/A") Or sedol = "" Then
                sedol = UCase("#--")
            End If
            If isin = UCase("#N/A N/A") Or isin = UCase("#N/A Invalid Security") Or isin = UCase("#N/A Field Not Applicable") Or _
                isin = UCase("#N/A") Or isin = "" Then
                isin = UCase("#--")
            End If
            If short_name = UCase("#N/A N/A") Or short_name = UCase("#N/A Invalid Security") Or short_name = UCase("#N/A Field Not Applicable") Or _
                short_name = UCase("#N/A") Or short_name = "" Then
                short_name = UCase("#--")
            End If
            If ccy = UCase("#N/A N/A") Or ccy = UCase("#N/A Invalid Security") Or ccy = UCase("#N/A Field Not Applicable") Or _
                ccy = UCase("#N/A") Or ccy = "" Then
                ccy = UCase("#--")
            End If
            If secdesc = UCase("#N/A N/A") Or secdesc = UCase("#N/A Invalid Security") Or secdesc = UCase("#N/A Field Not Applicable") Or _
                secdesc = UCase("#N/A") Or secdesc = "" Then
                secdesc = UCase("#--")
            End If
            If country = UCase("#N/A N/A") Or country = UCase("#N/A Invalid Security") Or country = UCase("#N/A Field Not Applicable") Or _
                country = UCase("#N/A") Or country = "" Then
                country = UCase("#--")
            End If
            If eq_fundccy = UCase("#N/A N/A") Or eq_fundccy = UCase("#N/A Invalid Security") Or eq_fundccy = UCase("#N/A Field Not Applicable") Or _
                eq_fundccy = UCase("#N/A") Or eq_fundccy = "" Then
                eq_fundccy = UCase("#--")
            End If
            If eq_sector = UCase("#N/A N/A") Or eq_sector = UCase("#N/A Invalid Security") Or eq_sector = UCase("#N/A Field Not Applicable") Or _
                eq_sector = UCase("#N/A") Or eq_sector = "" Then
                eq_sector = UCase("#--")
            End If
            If eq_industry = UCase("#N/A N/A") Or eq_industry = UCase("#N/A Invalid Security") Or eq_industry = UCase("#N/A Field Not Applicable") Or _
                eq_industry = UCase("#N/A") Or eq_industry = "" Then
                eq_industry = UCase("#--")
            End If
            If eq_subindustry = UCase("#N/A N/A") Or eq_subindustry = UCase("#N/A Invalid Security") Or eq_subindustry = UCase("#N/A Field Not Applicable") Or _
                eq_subindustry = UCase("#N/A") Or eq_subindustry = "" Then
                eq_subindustry = UCase("#--")
            End If
            If opt_maturity = UCase("#N/A N/A") Or opt_maturity = UCase("#N/A Invalid Security") Or opt_maturity = UCase("#N/A Field Not Applicable") Or _
                opt_maturity = UCase("#N/A") Or opt_maturity = "" Then
                opt_maturity = "1970-01-01"
            End If
            If opt_strike = UCase("#N/A N/A") Or opt_strike = UCase("#N/A Invalid Security") Or opt_strike = UCase("#N/A Field Not Applicable") Or _
                opt_strike = UCase("#N/A") Or opt_strike = "" Then
                opt_strike = -1
            End If
            If opt_valuepoint = UCase("#N/A N/A") Or opt_valuepoint = UCase("#N/A Invalid Security") Or UCase(opt_valuepoint) = UCase("#N/A Field Not Applicable") Or _
                opt_valuepoint = UCase("#N/A") Or opt_valuepoint = "" Then
                opt_valuepoint = -1
            End If
            If opt_underlying = UCase("#N/A N/A") Or opt_underlying = UCase("#N/A Invalid Security") Or opt_underlying = UCase("#N/A Field Not Applicable") Or _
                opt_underlying = UCase("#N/A") Or opt_underlying = "" Then
                opt_underlying = UCase("#--")
            End If
            If fut_underlying = UCase("#N/A N/A") Or fut_underlying = UCase("#N/A Invalid Security") Or fut_underlying = UCase("#N/A Field Not Applicable") Or _
                fut_underlying = UCase("#N/A") Or fut_underlying = "" Then
                fut_underlying = UCase("#--")
            End If
            If fut_valuepoint = UCase("#N/A N/A") Or fut_valuepoint = UCase("#N/A Invalid Security") Or fut_valuepoint = UCase("#N/A Field Not Applicable") Or _
                fut_valuepoint = UCase("#N/A") Or fut_valuepoint = "" Then
                fut_valuepoint = -1
            End If
            If fut_tickvalue = UCase("#N/A N/A") Or fut_tickvalue = UCase("#N/A Invalid Security") Or fut_tickvalue = UCase("#N/A Field Not Applicable") Or _
                fut_tickvalue = UCase("#N/A") Or fut_tickvalue = "" Then
                fut_tickvalue = -1
            End If
            If fut_ticksize = UCase("#N/A N/A") Or fut_ticksize = UCase("#N/A Invalid Security") Or fut_ticksize = UCase("#N/A Field Not Applicable") Or _
                fut_ticksize = UCase("#N/A") Or fut_ticksize = "" Then
                fut_ticksize = -1
            End If
            If fut_contractsize = UCase("#N/A N/A") Or fut_contractsize = UCase("#N/A Invalid Security") Or fut_contractsize = UCase("#N/A Field Not Applicable") Or _
                fut_contractsize = UCase("#N/A") Or fut_contractsize = "" Then
                fut_contractsize = -1
            End If
            If fut_firsttradedate = UCase("#N/A N/A") Or fut_firsttradedate = UCase("#N/A Invalid Security") Or fut_firsttradedate = UCase("#N/A Field Not Applicable") Or _
                fut_firsttradedate = UCase("#N/A") Or fut_firsttradedate = "" Then
                fut_firsttradedate = "1970-01-01"
            End If
            If fut_lasttradedate = UCase("#N/A N/A") Or fut_lasttradedate = UCase("#N/A Invalid Security") Or fut_lasttradedate = UCase("#N/A Field Not Applicable") Or _
                fut_lasttradedate = UCase("#N/A") Or fut_lasttradedate = "" Then
                fut_lasttradedate = "1970-01-01"
            End If
            
            strSQL = "INSERT INTO " & tblddbb & _
                        " (ticker, sectype, sedol, isin, short_name, ccy, secdesc, country, " & _
                        "eq_fundccy, eq_sector, eq_industry, eq_subindustry, " & _
                        "opt_maturity, opt_strike, opt_valuepoint, opt_underlying, " & _
                        "fut_underlying, fut_valuepoint, fut_tickvalue, fut_ticksize, " & _
                        "fut_contractsize, fut_firsttradedate, fut_lasttradedate, " & _
                        "isvalid) "
                        
            strSQL = strSQL & " VALUES ('" & _
                        ticker & "','" & sectype & "','" & sedol & "','" & isin & "','" & _
                        short_name & "','" & ccy & "','" & secdesc & "','" & _
                        country & "','" & eq_fundccy & "','" & eq_sector & "','" & _
                        eq_industry & "','" & eq_subindustry & "','" & opt_maturity & "','" & _
                        opt_strike & "','" & opt_valuepoint & "','" & opt_underlying & "','" & _
                        fut_underlying & "','" & fut_valuepoint & "','" & fut_tickvalue & "','" & _
                        fut_ticksize & "','" & fut_contractsize & "','" & fut_firsttradedate & "','" & _
                        fut_lasttradedate & "'," & isvalidasset & ")"
                        
            strSQL = strSQL & " ON DUPLICATE KEY UPDATE " & _
                        "ticker = '" & ticker & "', " & "sectype = '" & sectype & "', " & "sedol = '" & sedol & "', " & _
                        "isin = '" & isin & "', " & "short_name = '" & short_name & "', " & _
                        "ccy = '" & ccy & "', " & "secdesc = '" & secdesc & "', " & _
                        "country = '" & country & "', " & "eq_fundccy = '" & eq_fundccy & "', " & _
                        "eq_sector = '" & eq_sector & "', " & "eq_industry = '" & eq_industry & "', " & _
                        "eq_subindustry = '" & eq_subindustry & "', " & "opt_maturity = '" & opt_maturity & "', " & _
                        "opt_strike = '" & opt_strike & "', " & "opt_valuepoint = '" & opt_valuepoint & "', " & _
                        "opt_underlying = '" & opt_underlying & "', " & "fut_underlying = '" & fut_underlying & "', " & _
                        "fut_valuepoint = '" & fut_valuepoint & "', " & "fut_tickvalue = '" & fut_tickvalue & "', " & _
                        "fut_ticksize = '" & fut_ticksize & "', " & "fut_contractsize = '" & fut_contractsize & "', " & _
                        "fut_firsttradedate = '" & fut_firsttradedate & "', " & "fut_lasttradedate = '" & fut_lasttradedate & "', " & _
                        "isvalid = " & isvalidasset & ";"
            
            cmd.CommandText = strSQL
            cmd.Execute strSQL
            counter = counter + 1
            
            t = "31/07/2015"
            k = Sheets("Hist.").Cells(sht.rows.Count, "A").End(xlUp).Row + 1
            While t <= Date
                Sheets("Hist.").Cells(k, 1) = ticker
                Sheets("Hist.").Cells(k, 2) = sectype
                dd_str = CStr(Format(Day(t), "00"))
                mm_str = CStr(Format(Month(t), "00"))
                yy_str = CStr(Format(Year(t), "0000"))
                Sheets("Hist.").Cells(k, 3).NumberFormat = "@"
                Sheets("Hist.").Cells(k, 3) = yy_str & "-" & mm_str & "-" & dd_str
                Sheets("Hist.").Cells(k, 4) = "=BDH(A" & CStr(k) & "& "" "" & B" & CStr(k) & ",""" & strField & """,C" & CStr(k) & ", C" & CStr(k) & ")"
                Sheets("Hist.").Cells(k, 5) = "=BDH(A" & CStr(k) & "& "" "" & B" & CStr(k) & ",""VOLUME"",C" & CStr(k) & ", C" & CStr(k) & ")"
                k = k + 1
                t = DateAdd("d", 1, t)
            Wend
        Next i
        
        
        [hist_ticker] = hist_Tickers
        
        MsgBox ("A total of " & CStr(counter) & " new records successfully inserted into MySQL. " & vbNewLine & vbNewLine & _
                "IMPORTANT! Go to 'Hist tab' and insert historical prices before continue!")
                
    End If
    
    conn.Close
    Set conn = Nothing
    
End Sub

Sub blotter_autocomplete()
    
    Dim sht As Worksheet
    
    Application.ScreenUpdating = False
    
    Set sht = ActiveSheet
    
    lastrow = getLastRow
    
    Set rng = sht.Range("A10:O" & lastrow)
    For Each c In rng
        c.Value = UCase(Trim(c.Value))
    Next c
    
    For j = 11 To lastrow
        If sht.Cells(j, 1) = "" Then            ' ID
            sht.Cells(j, 1) = randomStr
        End If
        If sht.Cells(j, 2) = "" Then            ' Date
            dd_str = CStr(Format(Day(Date), "00"))
            mm_str = CStr(Format(Month(Date), "00"))
            yy_str = CStr(Format(Year(Date), "0000"))
            sht.Cells(j, 2).NumberFormat = "@"
            sht.Cells(j, 2) = yy_str & "-" & mm_str & "-" & dd_str
        End If
        If sht.Cells(j, 3) = "" Then            ' Date
            sht.Cells(j, 3) = Time
            sht.Cells(j, 3).NumberFormat = "hh:mm:ss"
        End If
        If sht.Cells(j, 8) = "" Then            ' Name
            sht.Cells(j, 8) = selectfromddbb("short_name", "assets", "where ticker = '" & sht.Cells(j, 7) & "'")
        End If
        If sht.Cells(j, 10) = "" Then            ' Status
            If sht.Cells(j, 11) = 0 Then
                sht.Cells(j, 10) = "--"
            Else
                sht.Cells(j, 10) = sht.Cells(j, 4) & "/" & sht.Cells(j, 5) & " " & _
                                    Format(sht.Cells(j, 11), "#,##0") & " " & sht.Cells(j, 7) & " @ " & Format(sht.Cells(j, 13), "#,##0.0000")
            End If
        End If
        If sht.Cells(j, 12) = "" Then            ' Currency
            sht.Cells(j, 12) = selectfromddbb("ccy", "assets", "where ticker = '" & sht.Cells(j, 7) & "'")
        End If
        If sht.Cells(j, 16) = "" Then            ' Remaining
            sht.Cells(j, 16) = sht.Cells(j, 6) - sht.Cells(j, 11)
        End If
    Next j

    blotter_formatCells
    Application.ScreenUpdating = True
End Sub

Sub assets_autocomplete()

    Dim sht As Worksheet
    
    Application.ScreenUpdating = False
    
    Set sht = ActiveSheet
    
    lastrow = getLastRow
    For j = 11 To lastrow
        If Not IsError(sht.Cells(j, 3)) Then
            If sht.Cells(j, 3) = "" Then            ' SEDOL
                sht.Cells(j, 3) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""ID_SEDOL1"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 4)) Then
            If sht.Cells(j, 4) = "" Then            ' ISIN
                sht.Cells(j, 4) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""ID_ISIN"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 5)) Then
            If sht.Cells(j, 5) = "" Then                ' NAME
                sht.Cells(j, 5) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""SHORT_NAME"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 6)) Then
            If sht.Cells(j, 6) = "" Then                ' CCY
                sht.Cells(j, 6) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""CRNCY"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 7)) Then
            If sht.Cells(j, 7) = "" Then                ' SEC. DESC.
                sht.Cells(j, 7) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""SECURITY_DES"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 8)) Then
            If sht.Cells(j, 8) = "" Then                ' COUNTRY
                sht.Cells(j, 8) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""COUNTRY_FULL_NAME"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 9)) Then
            If sht.Cells(j, 9) = "" Then                ' EQTY-FUND. CCY
                sht.Cells(j, 9) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""EQY_FUND_CRNCY"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 10)) Then
            If sht.Cells(j, 10) = "" Then               ' EQTY-SECTOR
                sht.Cells(j, 10) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""GICS_SECTOR_NAME"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 11)) Then
            If sht.Cells(j, 11) = "" Then               ' EQTY-INDUSTRY
                sht.Cells(j, 11) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""GICS_INDUSTRY_NAME"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 12)) Then
            If sht.Cells(j, 12) = "" Then               ' EQTY-SUBINDUSTRY
                sht.Cells(j, 12) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""GICS_SUB_INDUSTRY_NAME"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 13)) Then
            If sht.Cells(j, 13) = "" Then               ' OPT-MATURITY
                sht.Cells(j, 13) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""MATURITY"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 14)) Then
            If sht.Cells(j, 14) = "" Then               ' OPT-STRIKE
                sht.Cells(j, 14) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""OPT_STRIKE_PX"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 15)) Then
            If sht.Cells(j, 15) = "" Then               ' OPT-VALUE 1 PT
                sht.Cells(j, 15) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""OPT_VAL_PT"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 16)) Then
            If sht.Cells(j, 16) = "" Then               ' EQTY-UNDERLYING
                sht.Cells(j, 16) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""OPT_UNDL_TICKER"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 17)) Then
            If sht.Cells(j, 17) = "" Then               ' FUT-UNDERLYING
                sht.Cells(j, 17) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""UNDL_SPOT_TICKER"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 18)) Then
            If sht.Cells(j, 18) = "" Then               ' FUT-VAL 1 PT
                sht.Cells(j, 18) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""FUT_VAL_PT"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 19)) Then
            If sht.Cells(j, 19) = "" Then               ' FUT-TICK VALUE
                sht.Cells(j, 19) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""FUT_TICK_VAL"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 20)) Then
            If sht.Cells(j, 20) = "" Then               ' FUT-TICK SIZE
                sht.Cells(j, 20) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""FUT_TICK_SIZE"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 21)) Then
            If sht.Cells(j, 21) = "" Then               ' FUT-CONTRACT SIZE
                sht.Cells(j, 21) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""FUT_CONT_SIZE"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 22)) Then
            If sht.Cells(j, 22) = "" Then               ' FUT-FIRST TRADE DATE
                sht.Cells(j, 22) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""FUT_FIRST_TRADE_DT"")"
            End If
        End If
        If Not IsError(sht.Cells(j, 23)) Then
            If sht.Cells(j, 23) = "" Then               ' FUT-LAST TRADE DATE
                sht.Cells(j, 23) = "=BDP(A" & CStr(j) & "& "" "" & B" & CStr(j) & ", ""LAST_TRADEABLE_DT"")"
            End If
        End If
        If sht.Cells(j, 24) = "" Then                   ' IS-VALID
            sht.Cells(j, 24) = 1
        End If
    Next j
    
    Application.ScreenUpdating = True

End Sub

Public Function selectfromddbb(fldsddbb As String, tblddbb As String, Optional conditionddbb As String = "", Optional errDDBB As Variant)

 Dim conn As ADODB.Connection
    Dim SQLStr As String
    Dim rs As ADODB.Recordset
    
    Dim ipddbb As Variant
    Dim portddbb As Integer
    Dim nameddbb As Variant
    Dim userddbb As Variant
    Dim passddbb As Variant
    
    Dim sht As Worksheet
    
    ipddbb = ""
    portddbb = 3306
    nameddbb = ""
    userddbb = ""
    passddbb = ""
    
    If Not IsEmpty([mysql_server]) Then
        ipddbb = [mysql_server]
    End If
    If Not IsEmpty([mysql_port]) Then
        portddbb = [mysql_port]
    End If
    If Not IsEmpty([mysql_ddbb]) Then
        nameddbb = [mysql_ddbb]
    End If
    If Not IsEmpty([mysql_user]) Then
        userddbb = [mysql_user]
    End If
    If Not IsEmpty([mysql_pass]) Then
        passddbb = [mysql_pass]
    End If
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset

    If IsMissing(conditionddbb) Then
        conditionddbb = ""
    End If
    If IsMissing(errDDBB) Then
        errDDBB = UCase("#--")
    End If
    
    strSQL = "SELECT IFNULL( " & _
                "(SELECT " & fldsddbb & _
                " FROM " & tblddbb & _
                " " & conditionddbb & ")" & _
             ", '" & errDDBB & "');"
    
    conn.Open "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
                "SERVER=" & ipddbb & ";" & _
                "PORT=" & CStr(portddbb) & ";" & _
                "DATABASE=" & nameddbb & ";" & _
                "USER=" & userddbb & ";" & _
                "PASSWORD=" & passddbb & ";" & _
                "OPTION=3;"""
    
    rs.Open strSQL, conn, adOpenStatic
    
    '----------------------------------------------
    Dim myArray()
    
    selectfromddbb = "#--"
    
    Do While Not rs.EOF
    
        myArray = rs.GetRows()
    
        m = UBound(myArray, 2)
        n = UBound(myArray, 1)
                
        i = 0
        j = 0
        Do While i <= m ' Using For loop data are displayed
            Do While j <= n
                thisdata = myArray(j, i)
                If thisdata = "#N/A" Then
                    selectfromddbb = "#--"
                Else
                    selectfromddbb = thisdata
                End If
                j = j + 1
                thisdata = ""
            Loop
            i = i + 1
            j = 0
        Loop
        
    Loop
        '----------------------------------------------

    rs.Close
    Set rs = Nothing
    
    conn.Close
    Set conn = Nothing
End Function

Sub blotter_deleteTrades()
    
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
    Dim rng As Range
    Dim c As Range
    
    Set sht = ActiveSheet
    
    ipddbb = ""
    portddbb = 3306
    nameddbb = ""
    userddbb = ""
    passddbb = ""
    
    If Not IsEmpty([mysql_server]) Then
        ipddbb = [mysql_server]
    End If
    If Not IsEmpty([mysql_ddbb]) Then
        nameddbb = [mysql_ddbb]
    End If
    If Not IsEmpty([mysql_port]) Then
        portddbb = [mysql_port]
    End If
    If Not IsEmpty([mysql_user]) Then
        userddbb = [mysql_user]
    End If
    If Not IsEmpty([mysql_pass]) Then
        passddbb = [mysql_pass]
    End If
    
    lastrow = getLastRow
    
    Set rng = sht.Range("A10:O" & lastrow)
    For Each c In rng
        c.Value = UCase(Trim(c.Value))
    Next c

    msg = blotter_validateFields
    If msg <> "" Then
        MsgBox (msg)
        Exit Sub
    End If
    
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
    
    tblddbb = "fundtrades"
    If sht.Name = "BLOTTER-SWAP" Then
        tblddbb = "swaptrades"
    End If
    
    If lastrow > 10 Then
        counter = 0
        resp = MsgBox("DELETE from MySQL:" & vbCrLf & _
                      "<Yes> Delete" & vbCrLf & _
                      "<No> Cancel" & vbCrLf, _
                      vbYesNo, _
                      "Delete Trades")
        
        If resp = 7 Then
            Exit Sub
        End If
        
        For i = 11 To lastrow
            
            trd_id = sht.Cells(i, 1)                                                        ' ID
            
            strSQL = "DELETE FROM " & tblddbb & _
                        " WHERE trd_id = '" & trd_id & "';"
            
            cmd.CommandText = strSQL
            cmd.Execute strSQL
            
            Worksheets(sht.Name).rows(i).ClearContents
            counter = counter + 1
            
        Next i
        
        MsgBox ("A total of " & CStr(counter) & "  records successfully deleted from MySQL.")
        
    End If
    
    conn.Close
    Set conn = Nothing

    
    Application.ScreenUpdating = True
    
End Sub

Sub w2file(strFile_Path As String, line As String)
    Open strFile_Path For Append As #1
    Write #1, line
    Close #1
End Sub

Function bdhX(ticker As String, field As String, t0 As String, t1 As String) As Double

    On Error GoTo ErrorFound

    Dim conn As ADODB.Connection
    Dim SQLStr As String
    Dim rs As ADODB.Recordset
    
    Dim ipddbb As String
    Dim nameddbb As String
    Dim userddbb As String
    Dim passdd As String
    Dim tblddbb As String
    
    Dim amnt As Double
    Dim amnt0 As Double
    Dim px_t0_ccy As Double
    Dim px_avg_ccy As Double
    Dim px_avg_eur As Double
    Dim perf_cnt_ccy As Double
    Dim perf_eur As Double
    Dim xrate As Double
    Dim trdDate As String
    Dim trdType As String
    Dim trdPosition As String
    Dim trdFilled As Double
    Dim trdPriceCcy As Double
    Dim trdCcy As String
    Dim t_ini As String
    Dim firstDay As String
    Dim lastDay As String
    Dim prevClDay As String
    Dim firstDayM1 As String
    Dim tmpdate As String
    Dim v1p As Double
    
    Dim sht As Worksheet
    Set sht = ActiveSheet
    
    Dim strFile_Path As String
    
    Dim sngStart As Single, sngEnd As Single
    Dim sngElapsed As Single
    
    sngStart = Timer                               ' Get start time.
    
    ipddbb = ""
    portddbb = 3306
    nameddbb = ""
    userddbb = ""
    passddbb = ""
    
    field = UCase(field)
    
    If Not IsEmpty([mysql_server]) Then
        ipddbb = [mysql_server]
    End If
    If Not IsEmpty([mysql_ddbb]) Then
        nameddbb = [mysql_ddbb]
    End If
    If Not IsEmpty([mysql_port]) Then
        portddbb = [mysql_port]
    End If
    If Not IsEmpty([mysql_user]) Then
        userddbb = [mysql_user]
    End If
    If Not IsEmpty([mysql_pass]) Then
        passddbb = [mysql_pass]
    End If
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    tblddbb = "fundtrades"
    If sht.Name = "SWAP" Then
        tblddbb = "swaptrades"
    End If
    
    ccy = selectfromddbb("ccy", "assets", "WHERE ticker = '" & ticker & "'")
    
    t_ini = Format(CDate("31/07/2015"), "yyyy-mm-dd")
    firstDay = Format(CDate(t0), "yyyy-mm-dd")
    firstDayM1 = Format(CDate(t0) - 1, "yyyy-mm-dd")
    lastDay = Format(CDate(t1), "yyyy-mm-dd")
    
    If CDate(selectfromddbb("max(ha_date)", "histassets", "WHERE ha_ticker = '" & ticker & "'")) < CDate(lastDay) Then
        MsgBox ("Please, update historical prices in the DDBB before continue")
        GoTo ErrorFound
    End If
    If ccy <> "EUR" Then
        lastDayccy = selectfromddbb("max(ccy_date)", "histccy", "WHERE ccy_name = '" & ccy & "'")
        If CDate(lastDayccy) < CDate(lastDay) Then
            MsgBox ("Please, update prices (" & Format(lastDay, "yyyy-mm-dd") & ") and/or ccy (" & Format(lastDayccy, "yyyy-mm-dd") & ")")
            GoTo ErrorFound
        End If
    End If
    
    v1p = 1
    sectype = selectfromddbb("sectype", "assets", "WHERE ticker = '" & ticker & "'")
    If (sectype = "INDEX" And ticker <> "MCXP" And ticker <> "SCXP") Then
        v1p = selectfromddbb("fut_valuepoint", "assets", "WHERE ticker = '" & ticker & "'")
        firsttradedate = CDate(selectfromddbb("fut_firsttradedate", "assets", "WHERE ticker = '" & ticker & "'"))
        lasttradedate = CDate(selectfromddbb("fut_lasttradedate", "assets", "WHERE ticker = '" & ticker & "'"))
        If firsttradedate > CDate(firstDay) Then
            firstDay = CStr(Format(firsttradedate, "yyyy-mm-dd"))
            firstDayM1 = CStr(Format(firsttradedate - 1, "yyyy-mm-dd"))
        End If
        If lasttradedate < CDate(lastDay) Then
            lastDay = CStr(Format(lasttradedate, "yyyy-mm-dd"))
        End If
        If firstDay > lastDay Then
            GoTo maturity
        End If
    End If
    
    prevClDay = Format(CStr(CDate(lastDay) - 1), "yyyy-mm-dd")
    
    Dim myArray()
    
    '----------------------------------------------
    conn.Open "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
            "SERVER=" & ipddbb & ";" & _
            "PORT=3306;" & _
            "DATABASE=" & nameddbb & ";" & _
            "USER=" & userddbb & ";" & _
            "PASSWORD=" & passddbb & ";" & _
            "OPTION=3;"""
    
    '----------------------------------------------
    ' Portfolio between T0 - T1
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    conn.Open "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
        "SERVER=" & ipddbb & ";" & _
        "PORT=3306;" & _
        "DATABASE=" & nameddbb & ";" & _
        "USER=" & userddbb & ";" & _
        "PASSWORD=" & passddbb & ";" & _
        "OPTION=3;"""

    strSQL = "SELECT " & _
                "DATE_FORMAT(trd_date,'%Y-%m-%d'), " & _
                "trd_type, " & _
                "trd_position, " & _
                "trd_filled, " & _
                "trd_pricegross, " & _
                "trd_ccy " & _
            "FROM " & _
                tblddbb & " " & _
            "WHERE " & _
                "trd_ticker='" & ticker & "' " & _
                "AND trd_date>'" & t_ini & "' " & _
                "AND trd_date<='" & lastDay & "' " & _
            "ORDER BY " & _
                "trd_date ASC;"

    rs.Open strSQL, conn, adOpenStatic
    
    amnt = 0
    amnt0 = 0
    px_avg_ccy = 0
    px_avg_eur = 0
    perf_cnt_ccy = 0
    perf_eur = 0
    prev_close_ccy = 0
    prev_close_eur = 0
    
    max_count = 10
    tmp_count = max_count
    
    amnt = selectfromddbb("trd_filled", tblddbb, "WHERE trd_ticker = '" & ticker & "' AND trd_date='" & t_ini & "'", 0)
    
    If amnt <> 0 Then
    
        If selectfromddbb("trd_position", tblddbb, "WHERE trd_ticker = '" & ticker & "' AND trd_date='" & t_ini & "'", 0) = "SHORT" Then
            amnt = -amnt
        End If
    
        ' Get px and xrate the iniDate (31/07/2015)
        px_avg_ccy = -1
        xrate = -1
        tmpdate = t_ini
        While (px_avg_ccy = -1 Or xrate = -1)
            px_avg_ccy = selectfromddbb("trd_pricegross", tblddbb, "WHERE trd_ticker = '" & ticker & "' AND trd_date='" & tmpdate & "'", 0) * v1p
            If ccy <> "EUR" Then
                xrate = selectfromddbb("ccy_xrate", "histccy", "WHERE ccy_name = '" & ccy & "' AND ccy_date='" & tmpdate & "'", -1)
            Else
                xrate = 1
            End If
            t_ini = tmpdate
            tmpdate = Format(CDate(tmpdate) - 1, "yyyy-mm-dd")
            tmp_count = tmp_count - 1
            If tmp_count <= 0 Then
                GoTo ErrorNoPrices
            End If
        Wend
        tmp_count = max_count
        amnt0 = amnt
    
        If ccy = "GBP" Then
            px_avg_ccy = px_avg_ccy / 100
        End If
    
        px_avg_eur = px_avg_ccy / xrate
        
    End If
    
    Do While Not rs.EOF
        myArray = rs.GetRows()
        For j = 0 To UBound(myArray, 2)
            trdDate = Format(UCase(myArray(0, j)), "yyyy-mm-dd")
            trdType = UCase(myArray(1, j))
            trdPosition = UCase(myArray(2, j))
            trdFilled = Abs(myArray(3, j))
            trdPriceCcy = Abs(myArray(4, j) * v1p)
            trdCcy = UCase(myArray(5, j))
            If ccy <> "" Then
                If ccy <> trdCcy Then
                    MsgBox ("Please, check CCY. There is a mismatch between previous trades (" & ccy & ") and this trade(" & trdCcy & ")")
                    GoTo ErrorFound
                End If
            Else
                ccy = trdCcy
            End If
            If ccy = "GBP" Then
                trdPriceCcy = trdPriceCcy / 100
            End If
            
            xrate = 1
            If ccy <> "EUR" Then
                xrate = selectfromddbb("ccy_xrate", "histccy", "WHERE ccy_name = '" & ccy & "' AND ccy_date='" & trdDate & "'")
            End If
            
            inv_ccy_old = Abs(amnt * px_avg_ccy)
            inv_eur_old = Abs(amnt * px_avg_ccy / xrate)
            
            inv_ccy_new = Abs(trdFilled * trdPriceCcy)
            inv_eur_new = Abs(trdFilled * trdPriceCcy / xrate)
            
            If (CDate(trdDate) < CDate(firstDay)) Then
            
                If trdType = "BUY" Then
                    amnt = amnt + trdFilled
                Else
                    amnt = amnt - trdFilled
                End If
                
                amnt0 = amnt
                
            ElseIf CDate(trdDate) >= CDate(firstDay) And CDate(trdDate) <= CDate(lastDay) Then
                
                If trdType = "BUY" Then
                    amnt = amnt + trdFilled
                    perf_cnt_ccy = perf_cnt_ccy - inv_ccy_new
                    perf_eur = perf_eur - inv_eur_new
                Else
                    amnt = amnt - trdFilled
                    perf_cnt_ccy = perf_cnt_ccy + inv_ccy_new
                    perf_eur = perf_eur + inv_eur_new
                End If
                
            End If  ' END
            
            If amnt = 0 Then
                px_avg_ccy = 0
                px_avg_eur = 0
            ElseIf (amnt > 0 And trdType = "BUY") Or _
                (amnt < 0 And trdType = "SELL") Then
                px_avg_ccy = (inv_ccy_old + inv_ccy_new) / Abs(amnt)
                px_avg_eur = (inv_eur_old + inv_eur_new) / Abs(amnt)
            End If
            
            'Debug.Print j & ": " & inv_ccy_new & " - Perf: " & perf_cnt_ccy
            'Debug.Print j & px_ccy; ": " & px_avg_ccy & " - px_eur: " & px_avg_eur
            
        Next j
    Loop
    '----------------------------------------------

    rs.Close
    Set rs = Nothing
    
    conn.Close
    Set conn = Nothing
    
    ' Get px and xrate the firstDateM1
    px_first_ccy = -1
    xrate1 = -1
    tmpdate = firstDayM1
    While (px_first_ccy = -1 Or xrate1 = -1)
        px_first_ccy = selectfromddbb("ha_price", "histassets", "WHERE ha_ticker = '" & ticker & "' AND ha_date='" & tmpdate & "'", -1)
        If ccy <> "EUR" Then
            xrate1 = selectfromddbb("ccy_xrate", "histccy", "WHERE ccy_name = '" & ccy & "' AND ccy_date='" & tmpdate & "'", -1)
        Else
            xrate1 = 1
        End If
        firstDay = tmpdate
        tmpdate = Format(CDate(tmpdate) - 1, "yyyy-mm-dd")
        tmp_count = tmp_count - 1
        If tmp_count <= 0 Then
            GoTo ErrorNoPrices
        End If
    Wend
    tmp_count = max_count

    ' Get px and xrate the lastDate
    px_last_ccy = -1
    xrate2 = -1
    tmpdate = lastDay
    While px_last_ccy = -1
        px_last_ccy = selectfromddbb("ha_price", "histassets", "WHERE ha_ticker = '" & ticker & "' AND ha_date='" & tmpdate & "'", -1)
        If ccy <> "EUR" Then
            xrate2 = selectfromddbb("ccy_xrate", "histccy", "WHERE ccy_name = '" & ccy & "' AND ccy_date='" & tmpdate & "'", -1)
        Else
            xrate2 = 1
        End If
        lastDay = tmpdate
        tmpdate = Format(CDate(tmpdate) - 1, "yyyy-mm-dd")
        tmp_count = tmp_count - 1
        If tmp_count <= 0 Then
            GoTo ErrorNoPrices
        End If
    Wend
    tmp_count = max_count
    
    
    If (field = "DAY_CCY" Or field = "DAY_EUR") Then
    
        ' Get px and xrate the previouse closing date
        prev_close_ccy = -1
        xrate3 = -1
        tmpdate = prevClDay
        While prev_close_ccy = -1
            prev_close_ccy = selectfromddbb("ha_price", "histassets", "WHERE ha_ticker = '" & ticker & "' AND ha_date='" & tmpdate & "'", -1)
            If ccy <> "EUR" Then
                xrate3 = selectfromddbb("ccy_xrate", "histccy", "WHERE ccy_name = '" & ccy & "' AND ccy_date='" & tmpdate & "'", -1)
            Else
                xrate3 = 1
            End If
            prevClDay = tmpdate
            tmpdate = Format(CDate(tmpdate) - 1, "yyyy-mm-dd")
            tmp_count = tmp_count - 1
            If tmp_count <= 0 Then
                GoTo ErrorNoPrices
            End If
        Wend
        tmp_count = max_count
    
    Else
        prev_close_ccy = 1
        xrate3 = 1
    End If
    
    px_first_ccy = px_first_ccy * v1p
    px_last_ccy = px_last_ccy * v1p
    prev_close_ccy = prev_close_ccy * v1p
        
    If ccy = "GBP" Then
        px_first_ccy = px_first_ccy / 100
        px_last_ccy = px_last_ccy / 100
        prev_close_ccy = prev_close_ccy / 100
    End If
    
    px_first_eur = px_first_ccy / xrate1
    px_last_eur = px_last_ccy / xrate2
    prev_close_eur = prev_close_ccy / xrate3
    
    perf_cnt_ccy = -(amnt0 * px_first_ccy / xrate1) + perf_cnt_ccy / xrate1 + (amnt * px_last_ccy / xrate1)
    perf_eur = -(amnt0 * px_first_eur) + perf_eur + (amnt * px_last_eur)
    
    mval_ccy = Abs(amnt) * px_last_ccy
    mval_eur = Abs(amnt) * px_last_eur
    
    daily_ccy = (px_last_ccy / prev_close_ccy) - 1
    daily_eur = (px_last_eur / prev_close_eur) - 1
    
    If ccy = "GBP" Then
        px_first_ccy = px_first_ccy * 100
        px_last_ccy = px_last_ccy * 100
        px_avg_ccy = px_avg_ccy * 100
        prev_close_ccy = prev_close_ccy * 100
    End If
    
    
    sngEnd = Timer                                 ' Get end time.
    sngElapsed = Format((sngEnd - sngStart) * 1000, "##,##0.00") ' Elapsed time.

    'Debug.Print ticker & "|" & sngElapsed & ""
    
    
    strFile_Path = ThisWorkbook.Path & "\time.txt"
    'Call w2file(strFile_Path, ticker & "|" & sngElapsed & "")
        
    Select Case field
        Case "QTY"
            bdhX = FormatNumber(amnt, 0, vbTrue, vbTrue, vbTrue)
        Case "WEIGHT"
            bdhX = FormatNumber(mval_eur, 2, vbTrue, vbTrue, vbTrue)
            
        'CCY
        Case "PX_AVG_CCY"
            bdhX = FormatNumber(px_avg_ccy, 2, vbTrue, vbTrue, vbTrue)
        Case "PX_LAST_CCY"
            bdhX = FormatNumber(px_last_ccy, 2, vbTrue, vbTrue, vbTrue)
        Case "DAY_CCY"
            bdhX = FormatNumber(daily_ccy, 2, vbTrue, vbTrue, vbTrue)
        
        'EUR
        Case "PX_AVG_EUR"
            bdhX = FormatNumber(px_avg_eur, 2, vbTrue, vbTrue, vbTrue)
        Case "PX_LAST_EUR"
            bdhX = FormatNumber(px_last_eur, 2, vbTrue, vbTrue, vbTrue)
        Case "DAY_EUR"
            bdhX = FormatNumber(daily_eur, 2, vbTrue, vbTrue, vbTrue)
        
        Case "RTN_EUR"
            bdhX = FormatNumber(perf_eur, 2, vbTrue, vbTrue, vbTrue)
        Case "RTN_EUR_CNT_CCY"
            bdhX = FormatNumber(perf_cnt_ccy, 2, vbTrue, vbTrue, vbTrue)
        
        'Others
        Case Else
            bdhX = 0
    End Select

    Exit Function
    
ErrorFound:
    bdhX = "Error"
    Exit Function
    
ErrorNoPrices:
    bdhX = "No prices in DDBB"
    Exit Function

maturity:
    bdhX = 0
    Exit Function

End Function

Sub hist_get_tickers()

    Dim conn As ADODB.Connection
    Dim SQLStr As String
    Dim rs As ADODB.Recordset
    
    Dim ipddbb As String
    Dim nameddbb As String
    Dim userddbb As String
    Dim passdd As String
    Dim strFrom As String
    Dim strTo As String
    
    Dim sht As Worksheet
    Set sht = ActiveSheet
    
    ipddbb = ""
    portddbb = 3306
    nameddbb = ""
    userddbb = ""
    passddbb = ""
    
    field = UCase(field)
    
    If Not IsEmpty([mysql_server]) Then
        ipddbb = [mysql_server]
    End If
    If Not IsEmpty([mysql_ddbb]) Then
        nameddbb = [mysql_ddbb]
    End If
    If Not IsEmpty([mysql_port]) Then
        portddbb = [mysql_port]
    End If
    If Not IsEmpty([mysql_user]) Then
        userddbb = [mysql_user]
    End If
    If Not IsEmpty([mysql_pass]) Then
        passddbb = [mysql_pass]
    End If
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    strSelect = "*"
    strField = "PX_LAST"
    strFrom = CStr(Format(Date, "yyyy-mm-dd"))
    strTo = CStr(Format(Date, "yyyy-mm-dd"))
    
    If Not IsEmpty([hist_field]) Then
        strField = [hist_field]
    Else
        [hist_field] = strField
    End If
    If Not IsEmpty([hist_from]) Then
        strFrom = [hist_from]
    Else
        [hist_from] = strFrom
    End If
    If Not IsEmpty([hist_to]) Then
        strTo = [hist_to]
    Else
        [hist_to] = strTo
    End If
    If Not IsEmpty([hist_ticker]) Then
        strTicker = [hist_ticker]
    Else
        [hist_ticker] = strSelect
    End If
    
    '----------------------------------------------
    conn.Open "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
            "SERVER=" & ipddbb & ";" & _
            "PORT=3306;" & _
            "DATABASE=" & nameddbb & ";" & _
            "USER=" & userddbb & ";" & _
            "PASSWORD=" & passddbb & ";" & _
            "OPTION=3;"""
    
    If isvaliddate(strFrom) = False Or isvaliddate(strTo) = False Then
        MsgBox ("Please, enter a valid date with format yyyy-mm-dd")
        Exit Sub
    End If

    Dim daterng() As Single ' or whatever data type you wish to use
    k = 0
    For i = CDate(strFrom) To CDate(strTo)
        ReDim Preserve daterng(k)
        daterng(k) = i
        k = k + 1
    Next i

    tblddbb = "assets"
    If strTicker = "*" Or strTicker = "" Then
        
        strSQL = "SELECT " & _
                    "ticker, sectype " & _
                "FROM " & _
                    tblddbb & " " & _
                "WHERE " & _
                    "isvalid=1;"
    Else
        strSQL = "SELECT " & _
                    "ticker, sectype " & _
                "FROM " & _
                    tblddbb & " " & _
                "WHERE " & _
                    "ticker='" & strTicker & "';"
    End If
    
    rs.Open strSQL, conn, adOpenStatic
    
    Dim myArray()
    k = getLastRow + 1
    Do While Not rs.EOF
        myArray = rs.GetRows()
        For i = 0 To UBound(myArray, 2)
            If UBound(myArray, 1) = 1 Then
                For j = 0 To UBound(daterng, 1)
                    sht.Cells(k, 1) = myArray(0, i)
                    sht.Cells(k, 2) = myArray(1, i)
                    dd_str = CStr(Format(Day(daterng(j)), "00"))
                    mm_str = CStr(Format(Month(daterng(j)), "00"))
                    yy_str = CStr(Format(Year(daterng(j)), "0000"))
                    sht.Cells(k, 3).NumberFormat = "@"
                    sht.Cells(k, 3) = yy_str & "-" & mm_str & "-" & dd_str
                    sht.Cells(k, 4) = "=BDH(A" & CStr(k) & "& "" "" & B" & CStr(k) & ",""" & strField & """,C" & CStr(k) & ", C" & CStr(k) & ")"
                    sht.Cells(k, 5) = "=BDH(A" & CStr(k) & "& "" "" & B" & CStr(k) & ",""VOLUME"",C" & CStr(k) & ", C" & CStr(k) & ")"
                    k = k + 1
                Next j
            Else
                sht.Cells(k, 1) = myArray(0, i)
                sht.Cells(k, 3) = Format(myArray(1, i), "dd/mm/yyyy")
                sht.Cells(k, 4) = myArray(2, i)
                sht.Cells(k, 5) = myArray(3, i)
                k = k + 1
            End If
        Next i
    Loop
    '----------------------------------------------

    rs.Close
    Set rs = Nothing
    
    conn.Close
    Set conn = Nothing

End Sub


Sub hist_save_in_MySQL()

    Dim conn As ADODB.Connection
    Dim SQLStr As String
    Dim rs As ADODB.Recordset
    
    Dim ipddbb As String
    Dim nameddbb As String
    Dim userddbb As String
    Dim passdd As String
    Dim strFrom As String
    Dim strTo As String
    
    Dim sht As Worksheet
    Set sht = ActiveSheet
    
    ipddbb = ""
    portddbb = 3306
    nameddbb = ""
    userddbb = ""
    passddbb = ""
    
    field = UCase(field)
    
    If Not IsEmpty([mysql_server]) Then
        ipddbb = [mysql_server]
    End If
    If Not IsEmpty([mysql_ddbb]) Then
        nameddbb = [mysql_ddbb]
    End If
    If Not IsEmpty([mysql_port]) Then
        portddbb = [mysql_port]
    End If
    If Not IsEmpty([mysql_user]) Then
        userddbb = [mysql_user]
    End If
    If Not IsEmpty([mysql_pass]) Then
        passddbb = [mysql_pass]
    End If
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
        
    conn.Open "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
                "SERVER=" & ipddbb & ";" & _
                "PORT=" & CStr(portddbb) & ";" & _
                "DATABASE=" & nameddbb & ";" & _
                "USER=" & userddbb & ";" & _
                "PASSWORD=" & passddbb & ";" & _
                "OPTION=3;"""
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    
    lastrow = getLastRow
    Set rng = sht.Range("A10:B" & lastrow)
    For Each c In rng
        c.Value = UCase(Trim(c.Value))
    Next c
    Set rng = Nothing
    
    Set rng = sht.Range("D10:E" & lastrow)
    For Each c In rng
        If InStr(c.Value, "#N/A N/A") = 1 Then
            c.Value = -1
        End If
    Next c

    k = 11
    j = 11
    While k < lastrow
        If InStr(sht.Cells(j, 4), "#N/A Invalid Security") = 1 Then
            ticker = sht.Cells(j, 1).Value
            strSQL = "UPDATE assets SET isvalid=0 WHERE ticker = '" & ticker & "';"
            cmd.CommandText = strSQL
            cmd.Execute strSQL
            sht.rows(j).Delete
        Else
            j = j + 1
        End If
        k = k + 1
    Wend
    
    lastrow = getLastRow
    
    msg = hist_validateFields
    If msg <> "" Then
        MsgBox (msg)
        Exit Sub
    End If
    
    If lastrow > 10 Then
        counter = 0
        resp = MsgBox("Insert into MySQL:" & vbCrLf & _
                      "<Yes> Save" & vbCrLf & _
                      "<No> Abort" & vbCrLf, _
                      vbYesNo, _
                      "Historical data to MySQL")
        
        If resp = 7 Then
            Exit Sub
        End If
        
        tblddbb = "histassets"
        
        For i = 11 To lastrow
            
            hist_ticker = Replace(sht.Cells(i, 1), "'", "\'")                               ' Ticker
            hist_type = Replace(sht.Cells(i, 2), "'", "\'")                                 ' Asset type
            hist_date = Replace(sht.Cells(i, 3), "'", "\'")                                 ' Date
            hist_px = Replace(Format(CDbl(sht.Cells(i, 4)), "#0.0000"), ",", ".")           ' Price
            hist_vol = Replace(Format(CDbl(sht.Cells(i, 5)), "#0.0000"), ",", ".")          ' Volume
            
            strSQL = "INSERT INTO " & tblddbb & _
                        " (ha_ticker, ha_date, ha_price, ha_volume) "
                        
            strSQL = strSQL & " VALUES (" & _
                        "'" & hist_ticker & "'," & _
                        "'" & hist_date & "'," & _
                        " " & hist_px & "," & _
                        " " & hist_vol & ")"
                        
            strSQL = strSQL & " ON DUPLICATE KEY UPDATE " & _
                        "ha_ticker = '" & hist_ticker & "', " & _
                        "ha_date = '" & hist_date & "', " & _
                        "ha_price = " & hist_px & ", " & _
                        "ha_volume = " & hist_vol & ";"
            
            cmd.CommandText = strSQL
            cmd.Execute strSQL
            counter = counter + 1
        Next i
        
        MsgBox ("A total of " & CStr(counter) & " new records successfully inserted into MySQL.")
        
    End If
    
    conn.Close
    Set conn = Nothing


End Sub
