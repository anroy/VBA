Sub LoopThroughFiles()
    Dim strFile As String
    Dim strPath As String
    Dim colFiles As New Collection
    Dim i As Integer
   
    strPath = ActiveWorkbook.Path + "\"
    strFile = Dir(strPath)
   
    While strFile <> ""
        colFiles.Add strFile
        strFile = Dir
    Wend
   
    'List filenames in Column A of the active sheet
    If colFiles.Count > 0 Then
        For i = 1 To colFiles.Count
            ActiveSheet.Cells(i, 1).Value = colFiles(i)
        Next i
    End If
   
End Sub

