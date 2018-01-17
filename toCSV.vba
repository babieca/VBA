Sub trades_to_csv_file()
    
    Dim ColNum As Integer
    Dim Line As String
    Dim LineValues() As Variant
    Dim OutputFileNum As Integer
    Dim PathName As String
    Dim RowNum As Integer
    
    Dim ActSheet As Worksheet
    Dim SelRange As Range
    
    Dim sFile As String
    
    Set ActSheet = ActiveSheet
    Set SelRange = Selection
    
    'PathName = GetFolder("")
    PathName = Application.ActiveWorkbook.Path
    
    If Not Right(PathName, 1) = "\" Then
        PathName = PathName & "\"
    End If
    
    sFile = ActiveWorkbook.Name
    
    If InStr(sFile, ".") Then
        sFile = Left(sFile, InStr(sFile, ".") - 1)
    End If
    
    OutputFileNum = FreeFile
    
    Open PathName & sFile & ".csv" For Output Lock Write As #OutputFileNum
    
    ReDim LineValues(1 To SelRange.Columns.Count)

    For RowNum = 1 To SelRange.Rows.Count
        For ColNum = 1 To SelRange.Columns.Count
            this_value = SelRange(RowNum, ColNum).Value
            If IsNumeric(this_value) Then
                this_value = Replace(this_value, ",", "")
            ElseIf IsDate(this_value) Then
                this_value = Format(this_value, "mm/dd/yyyy")
            End If
            LineValues(ColNum) = this_value
        Next
        Line = Join(LineValues, "|")
        Print #OutputFileNum, Line
    Next

    Close OutputFileNum
    
    MsgBox "The document was saved successfully!  ;-)"
    
End Sub

Function GetFolder(strPath As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function


