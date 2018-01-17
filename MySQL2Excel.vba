Function MySQLtoExcel(fldsddbb As String, tblddbb As String, Optional conditionddbb As String = "", Optional errDDBB As Variant)

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
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset

    If IsMissing(conditionddbb) Then
        conditionddbb = ""
    End If
    If IsMissing(errDDBB) Then
        errDDBB = "#--"
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
    
    MySQLtoExcel = "#N/A"
    
    Do While Not rs.EOF
    
        myArray = rs.GetRows()
    
        kolumner = UBound(myArray, 1)
        rader = UBound(myArray, 2)
        
        For i = 0 To kolumner ' Using For loop data are displayed
            'Range("a5").Offset(0, K).Value = rs.Fields(K).Name
            For j = 0 To rader
                'Range("A5").Offset(R + 1, i).Value = myArray(i, j)
                If myArray(i, j) = "#N/A" Then
                    MySQLtoExcel = "--"
                Else
                    If Val(myArray(i, j)) = 0 And Not myArray(i, j) = "0.000000" Then
                        MySQLtoExcel = myArray(i, j)
                    Else
                        MySQLtoExcel = Val(myArray(i, j))
                    End If
                End If
            Next
        Next
        
    Loop
        '----------------------------------------------

    rs.Close
    Set rs = Nothing
    
    conn.Close
    Set conn = Nothing
    
End Function
