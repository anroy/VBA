Sub ProcessAllWorkbooks()

    Dim thisBook As String
    Dim thisPath As String
    thisBook = ActiveWorkbook.Name
    thisPath = ActiveWorkbook.Path + "\"

    Dim strFile As String
    Dim strFullFile As String
    Dim processedFiles As New Collection
    Dim i As Integer

    strFile = Dir(thisPath)

    While strFile <> ""
        If strFile <> thisBook Then

            strFullFile = thisPath + strFile
            Workbooks.Open Filename:=strFullFile
            Application.Run "Scripts.xls!CityStripSpaces"
            ActiveWorkbook.Save
            ActiveWorkbook.Close

            processedFiles.Add strFullFile
        End If
        strFile = Dir
    Wend

    'List processed filenames in Column A of the active sheet
    If processedFiles.Count > 0 Then
        For i = 1 To processedFiles.Count
            ActiveSheet.Cells(i, 1).Value = processedFiles(i)
        Next i
    End If

End Sub

