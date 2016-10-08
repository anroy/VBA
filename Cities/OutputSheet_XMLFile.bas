Sub OutputSheet_XMLFile()
    'Created 2014/10/11 by Arka
    If MsgBox("Export XML file?", vbYesNo) = vbNo Then Exit Sub
    
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    'objStream.Position = 0
    objStream.Type = 2
    objStream.Open

    'XML declaration
    objStream.writetext "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & Chr(10)

    'Root element opening tag with namespace attribute
    objStream.writetext "<netpeopleFAQ xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" & Chr(10)

    Dim rowCnt As Long, colCnt As Long
    'rowInd and colInd are 1-indexed
    Dim rowInd As Long, colInd As Long
    'first row contains headings
    Const startRow = 2
    'first column contains command button
    Const startCol = 2
    Dim tagName As String

    colCnt = ActiveSheet.Columns.Count
    rowCnt = ActiveSheet.Rows.Count

    For rowInd = startRow To rowCnt
        If Trim(Cells(rowInd, startCol).Value) = "" Then Exit For

        'article opening tag
        objStream.writetext "<article>" & Chr(10)

        For colInd = startCol To colCnt
            If Trim(Cells(rowInd, colInd).Value) = "" Then Exit For
            tagName = Cells(1, colInd)
            objStream.writetext CreateTag(1, tagName, Trim(Cells(rowInd, colInd).Value)) & Chr(10)
        Next colInd
        'For colInd

        'article closing tag
        objStream.writetext "</article>" & Chr(10)

    Next rowInd
    'For rowInd

    'Root element closing tag
    objStream.writetext "</netpeopleFAQ>" & Chr(10)

    outFileName = ActiveWorkbook.Path + "\" + Replace(ActiveWorkbook.Name, ".xlsm", "") + ".xml"
    objStream.SaveToFile outFileName, 2
End Sub


Function CreateTag(tabCnt As Integer, tagName As String, tagContent As String) As String
    tabCat = ""
    
    For i = 1 To tabCnt
        tabCat = tabCat & Chr(9)
    Next i
    
    CreateTag = tabCat & "<" & tagName & ">" & tagContent & "</" & tagName & ">"
End Function
