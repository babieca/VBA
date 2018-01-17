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

Function callpy(strSQL As String) As Variant

    Dim pythonExe As String
    Dim pythonScript As String
    Dim execStr As String, resp As String
    Dim response As Variant
    
    pythonExe = ".\python.exe"
    pythonScript = "emsx.v1.py"
    'pythonScript = ActiveWorkbook.Path & "\emsx.v1.py"
    
    execStr = pythonExe & " " & pythonScript & " " & strSQL
    
    response = ShellRun(execStr)
    
    'callpy = response
    callpy = Val(response)
    
End Function


Function callpyhide(strSQL As String) As Variant

    Dim pythonExe As String
    Dim pythonScript As String
    Dim execStr As String, launch As String
    Dim strOutput As Variant
    Dim outputFilePath As String
    
    pythonExe = ".\python.exe"
    pythonScript = ".\closing-books.py"
    outputFilePath = ".\output.txt"
    
    'pythonScript = Application.ActiveWorkbook.Path & ".\emsx.v1.py"
    'outputFilePath = Application.ActiveWorkbook.Path & "\output.txt"
    
    execStr = pythonExe & " " & pythonScript & " -b """ & Cells("C8").Value & """"

    launch = "cmd.exe /c " & execStr & " > " & outputFilePath
    With CreateObject("WScript.Shell")
        ' Pass 0 as the second parameter to hide the window...
        .Run launch, 0, True
    End With
    
    ' Read the output and remove the file when done...
    With CreateObject("Scripting.FileSystemObject")
    
        strOutput = .OpenTextFile(outputFilePath).ReadAll()
        .DeleteFile outputFilePath
    
    End With
    
    callpyhide = Val(strOutput)
    
    
End Function



