Attribute VB_Name = "XL_to_XML"
Sub MakeXML()
' create an XML file from an Excel table
Dim MyRow As Integer, MyCol As Integer, Temp As String, YesNo As Variant, DefFolder As String
Dim XMLFileName As String, XMLRecSetName As String, MyLF As String, RTC1 As Integer
Dim RangeOne As String, RangeTwo As String, Tt As String, FldName(99) As String

MyLF = Chr(10) & Chr(13)    ' a line feed command
DefFolder = "C:\"   'change this to the location of saved XML files

YesNo = MsgBox("This procedure requires the following data:" & MyLF _
 & "1 A filename for the XML file" & MyLF _
 & "2 A groupname for an XML record" & MyLF _
 & "3 A cellrange containing fieldnames (col titles)" & MyLF _
 & "4 A cellrange containing the data table" & MyLF _
 & "Are you ready to proceed?", vbQuestion + vbYesNo, "MakeXML CiM")
 
If YesNo = vbNo Then
 Debug.Print "User aborted with 'No'"
 Exit Sub
End If

XMLFileName = FillSpaces(InputBox("1. Enter the name of the XML file:", "MakeXML CiM", "xl_xml_data"))
If Right(XMLFileName, 4) <> ".xml" Then
 XMLFileName = XMLFileName & ".xml"
End If

XMLRecSetName = FillSpaces(InputBox("2. Enter an identifying name of a record:", "MakeXML CiM", "record"))

RangeOne = InputBox("3. Enter the range of cells containing the field names (or column titles):", "MakeXML CiM", "A3:D3")
If MyRng(RangeOne, 1) <> MyRng(RangeOne, 2) Then
  MsgBox "Error: names must be on a single row" & MyLF & "Procedure STOPPED", vbOKOnly + vbCritical, "MakeXML CiM"
  Exit Sub
End If
MyRow = MyRng(RangeOne, 1)
For MyCol = MyRng(RangeOne, 3) To MyRng(RangeOne, 4)
 If Len(Cells(MyRow, MyCol).Value) = 0 Then
  MsgBox "Error: names range contains blank cell" & MyLF & "Procedure STOPPED", vbOKOnly + vbCritical, "MakeXML CiM"
  Exit Sub
 End If
 FldName(MyCol - MyRng(RangeOne, 3)) = FillSpaces(Cells(MyRow, MyCol).Value)
Next MyCol

RangeTwo = InputBox("4. Enter the range of cells containing the data table:", "MakeXML CiM", "A4:D8")
If MyRng(RangeOne, 4) - MyRng(RangeOne, 3) <> MyRng(RangeTwo, 4) - MyRng(RangeTwo, 3) Then
  MsgBox "Error: number of field names <> data columns" & MyLF & "Procedure STOPPED", vbOKOnly + vbCritical, "MakeXML CiM"
  Exit Sub
End If
RTC1 = MyRng(RangeTwo, 3)

If InStr(1, XMLFileName, ":\") = 0 Then
 XMLFileName = DefFolder & XMLFileName
End If

Open XMLFileName For Output As #1
Print #1, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "ISO-8859-1" & Chr(34) & "?>"
Print #1, "<meadinkent>"

For MyRow = MyRng(RangeTwo, 1) To MyRng(RangeTwo, 2)
Print #1, "<" & XMLRecSetName & ">"
  For MyCol = RTC1 To MyRng(RangeTwo, 4)
  ' the next line uses the FormChk function to format dates and numbers
     Print #1, "<" & FldName(MyCol - RTC1) & ">" & RemoveAmpersands(FormChk(MyRow, MyCol)) & "</" & FldName(MyCol - RTC1) & ">"
  ' the next line does not apply any formatting
  '  Print #1, "<" & FldName(MyCol - RTC1) & ">" & RemoveAmpersands(Cells(MyRow, MyCol).Value) & "</" & FldName(MyCol - RTC1) & ">"
    Next MyCol
 Print #1, "</" & XMLRecSetName & ">"

Next MyRow
Print #1, "</meadinkent>"
Close #1
MsgBox XMLFileName & " created." & MyLF & "Process finished", vbOKOnly + vbInformation, "MakeXML CiM"
Debug.Print XMLFileName & " saved"
End Sub
Function MyRng(MyRangeAsText As String, MyItem As Integer) As Integer
' analyse a range, where MyItem represents 1=TR, 2=BR, 3=LHC, 4=RHC

Dim UserRange As Range
Set UserRange = Range(MyRangeAsText)
Select Case MyItem
 Case 1
 MyRng = UserRange.Row
 Case 2
 MyRng = UserRange.Row + UserRange.Rows.Count - 1
 Case 3
 MyRng = UserRange.Column
 Case 4
 MyRng = UserRange.Columns(UserRange.Columns.Count).Column
End Select
Exit Function

End Function
Function FillSpaces(AnyStr As String) As String
' remove any spaces and replace with underscore character
Dim MyPos As Integer
MyPos = InStr(1, AnyStr, " ")
Do While MyPos > 0
 Mid(AnyStr, MyPos, 1) = "_"
 MyPos = InStr(1, AnyStr, " ")
Loop
FillSpaces = LCase(AnyStr)
End Function

Function FormChk(RowNum As Integer, ColNum As Integer) As String
' formats numeric and date cell values to comma 000's and DD MMM YY
FormChk = Cells(RowNum, ColNum).Value
If IsNumeric(Cells(RowNum, ColNum).Value) Then
 FormChk = Format(Cells(RowNum, ColNum).Value, "#,##0 ;(#,##0)")
End If
If IsDate(Cells(RowNum, ColNum).Value) Then
 FormChk = Format(Cells(RowNum, ColNum).Value, "dd mmm yy")
End If
End Function

Function RemoveAmpersands(AnyStr As String) As String
Dim MyPos As Integer
' replace Ampersands (&) with plus symbols (+)

MyPos = InStr(1, AnyStr, "&")
Do While MyPos > 0
 Mid(AnyStr, MyPos, 1) = "+"
 MyPos = InStr(1, AnyStr, "&")
Loop
RemoveAmpersands = AnyStr
End Function
