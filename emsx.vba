Option Explicit

Public Function ShellRun(sCmd As String) As String
    ' https://stackoverflow.com/questions/2784367/capture-output-value-from-a-shell-command-in-vba

    'Run a shell command, returning the output as a string'

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command'
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    Set oOutput = oExec.stdout

    'handle the results as they are written to and read from the StdOut object'
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend

    ShellRun = s

End Function

Public Function emsx_report()

    Dim pythonExe As String
    Dim pythonScript As String
    Dim execStr As String, launch As String
    Dim strOutput As Variant
    Dim outputFilePath As String
    Dim daysback As Integer
    
    daysback = Date - ActiveWorkbook.Worksheets("BLOTTER-FUND").Cells(5, 2).Value
    
    pythonExe = ".\python.exe"
    pythonScript = ".\emsx.py"
    outputFilePath = ".\output.txt"
    
    execStr = pythonExe & " " & pythonScript & " -b """ & _
                    ActiveWorkbook.Worksheets("BLOTTER-FUND").Cells(4, 2).Value & """ " & daysback

    launch = "cmd.exe /c " & execStr '& " > " & outputFilePath
    With CreateObject("WScript.Shell")
        ' Pass 0 as the second parameter to hide the window...
        .Run launch, 0, True
    End With
    
    ' Read the output and remove the file when done...
    With CreateObject("Scripting.FileSystemObject")
    
        strOutput = .OpenTextFile(outputFilePath).ReadAll()
        .DeleteFile outputFilePath
    
    End With

    strOutput = Trim(CStr(strOutput))
    strOutput = Replace(strOutput, vbLf, "")
    strOutput = Replace(strOutput, vbCr, "")
    strOutput = Replace(strOutput, vbCrLf, "")
    strOutput = Replace(strOutput, vbTab, "")
    
    If strOutput = "{}" Then
        MsgBox ("Python output: " & strOutput)
    Else
        ReadJSON (strOutput)
    End If

End Function

Public Function ReadJSON(content As String)
 
    Dim root As Object
    Dim rootKeys() As String
    Dim keys() As String
    Dim i As Integer
    Dim obj As Object
    Dim prop As Variant
    
    content = Replace(content, vbCrLf, "")
    content = Replace(content, vbTab, "")
 
    JsonParser.InitScriptEngine
 
    Set root = JsonParser.DecodeJsonString(content)
  
    rootKeys = JsonParser.GetKeys(root)
    
    For i = 0 To UBound(rootKeys)
    
        'Debug.Print rootKeys(i)
        
        'If JsonParser.GetPropertyType(root, rootKeys(i)) = jptValue Then
        '    prop = JsonParser.GetProperty(root, rootKeys(i))
        '    Debug.Print prop
        'Else
            Set obj = JsonParser.GetObjectProperty(root, rootKeys(i))
            RecurseProps obj, 2
        'End If
        
    Next i
 
End Function
 
 
Private Function RecurseProps(obj As Object, Optional Indent As Integer = 0) As Object
    Dim nextObject As Object
    Dim propValue As Variant
    Dim keys() As String
    Dim i As Integer
    Dim k As Integer
    Dim lr As Integer
    Dim broker As Variant
    Dim namesht As String
    
    keys = JsonParser.GetKeys(obj)
    
    Dim key_broker As Integer
    
    For k = 0 To UBound(keys)
        If keys(k) = "broker" Then
            broker = JsonParser.GetProperty(obj, keys(k))
            If broker = "UBSX" Then
                namesht = "BLOTTER-SWAP"
            Else
                namesht = "BLOTTER-FUND"
            End If
            lr = getLastRow(namesht) + 1
            Exit For
        End If
    Next k
    
    For i = 0 To UBound(keys)
        
        broker = JsonParser.GetProperty(obj, keys(k))
        
        If JsonParser.GetPropertyType(obj, keys(i)) = jptValue Then
            
            propValue = JsonParser.GetProperty(obj, keys(i))
            'Debug.Print Space(Indent) & keys(i) & ": " & propValue
            
            If keys(i) = "ticker" Then
                ActiveWorkbook.Sheets(namesht).Cells(lr, 7).Value = propValue
                If namesht = "BLOTTER-FUND" And _
                    (Right(propValue, 3) = " LN" Or Right(propValue, 3) = " PL") Then
                    ActiveWorkbook.Sheets(namesht).Cells(lr, 15).Value = "AS CFD"
                End If
            ElseIf keys(i) = "price" Then
                ActiveWorkbook.Sheets(namesht).Cells(lr, 13).Value = propValue
            ElseIf keys(i) = "broker" Then
                ActiveWorkbook.Sheets(namesht).Cells(lr, 14).Value = propValue
            ElseIf keys(i) = "shares" Then
                ActiveWorkbook.Sheets(namesht).Cells(lr, 6).Value = propValue
                ActiveWorkbook.Sheets(namesht).Cells(lr, 11).Value = propValue
            ElseIf keys(i) = "date" Then
                ActiveWorkbook.Sheets(namesht).Cells(lr, 2).Value = propValue
                ActiveWorkbook.Sheets(namesht).Cells(lr, 3) = Time
                ActiveWorkbook.Sheets(namesht).Cells(lr, 3).NumberFormat = "hh:mm:ss"
            ElseIf keys(i) = "type" Then
                ActiveWorkbook.Sheets(namesht).Cells(lr, 4).Value = propValue
            ElseIf keys(i) = "id" Then
                ActiveWorkbook.Sheets(namesht).Cells(lr, 1).Value = propValue
            End If
        'Else
        '    Set nextObject = JsonParser.GetObjectProperty(obj, keys(i))
        '    Debug.Print Space(Indent) & keys(i)
        '    RecurseProps nextObject, Indent + 2
        End If
    
    Next i
    
End Function

Public Function getLastRow(namesht As String)

    Dim lastrow As Integer
    Dim lastRowF As Integer

    lastrow = Worksheets(namesht).Cells(Worksheets(namesht).rows.Count, "A").End(xlUp).Row
    lastRowF = Worksheets(namesht).Cells(Worksheets(namesht).rows.Count, "F").End(xlUp).Row
    If lastRowF > lastrow Then
        lastrow = lastRowF
    End If
    If lastrow <= 10 Then
        lastrow = 10
    End If
    getLastRow = lastrow
End Function

Sub clearfundswap()

    Dim lastrow As Integer
    lastrow = getLastRow("BLOTTER-FUND")
    
    If lastrow > 10 Then
        Sheets("BLOTTER-FUND").Range("A11:P" & lastrow).ClearContents
    End If
    
    lastrow = getLastRow("BLOTTER-SWAP")
    If lastrow > 10 Then
        Sheets("BLOTTER-SWAP").Range("A11:P" & lastrow).ClearContents
    End If
End Sub

