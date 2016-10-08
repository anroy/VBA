Public Function ExportToXML(FullPath As String, RowName _
  As String) As Boolean

'PURPOSE: EXPORTS AN EXCEL SPREADSHEET TO XML
'PARAMETERS: FullPath: Full Path of File to Export Sheet to
'             RowName: XML Attribute Name to give to each row

'RETURNS: True if Successful, false otherwise

'EXAMPLE: ExportToXML "C:\mysheet.xml", "Employee"

'NOTES:
'This function has the following quirks and limitations.
'If you find that they are not consistent with the behavior
'you desire for your solution, you should be able to
'modify the code without too much difficulty

'       1) Designed to be used inside Excel as a macro
'        not with VB.  If you want to use from VB
'        Add code to use Excel Object model
'
'       2) This snippet works with the
'          the first worksheet in the workbook.
'          If you want to make this a variable,
'          You can change the code to add the worksheet
'          Number as a parameter.
'
'       3) This code uses the worksheet name as the top-level
'          XML attribute.
'
'       4) The first row of the sheet is assumed to contain the
'          attribute (column) names, while the following rows
'          are assumed to contained the data values
'
'       5) No data for blank cells are written to the
'          XML file.
'
'       6) The CDATA attribute is included with each value
'
'       7) The function assumes that the first column of
'          each row in the sheet has a value.  If it finds a
'          blank first column it exits.  This is in order
'          to prevent it from printing blank row
'******************************************************

On Error GoTo ErrorHandler


Dim colIndex As Integer
Dim rwIndex As Integer
Dim asCols() As String
Dim oWorkSheet As Worksheet
Dim sName As String
Dim lCols As Long, lRows As Long
Dim iFileNum As Integer


Set oWorkSheet = ThisWorkbook.Worksheets(1)
sName = oWorkSheet.Name
lCols = oWorkSheet.Columns.Count
lRows = oWorkSheet.Rows.Count


ReDim asCols(lCols) As String

iFileNum = FreeFile
Open FullPath For Output As #iFileNum

For i = 0 To lCols - 1
    'Assumes no blank column names
    If Trim(Cells(1, i + 1).Value) = "" Then Exit For
    asCols(i) = Cells(1, i + 1).Value
Next i

If i = 0 Then GoTo ErrorHandler
lCols = i

Print #iFileNum, "<?xml version=""1.0""?>"
Print #iFileNum, "<" & sName & ">"
For i = 2 To lRows
If Trim(Cells(i, 1).Value) = "" Then Exit For
Print #iFileNum, "<" & RowName & ">"
  
    For j = 1 To lCols
        
        If Trim(Cells(i, j).Value) <> "" Then
           Print #iFileNum, "  <" & asCols(j - 1) & "><![CDATA[";
           Print #iFileNum, Trim(Cells(i, j).Value);
           Print #iFileNum, "]]></" & asCols(j - 1) & ">"
           DoEvents 'OPTIONAL
        End If
    Next j
    Print #iFileNum, " </" & RowName & ">"
Next i

Print #iFileNum, "</" & sName & ">"
ExportToXML = True
ErrorHandler:
If iFileNum > 0 Then Close #iFileNum
Exit Function
End Function