Sub OutputSheet()

'
' Created 2010/4/22 Arka
' Updated 2010/6/3 Arka

        Dim myRecord As Range
        Dim myField As Range
        Dim nFileNum As Long
        
        Dim wb As Workbook
        Set wb = ActiveWorkbook

        nFileNum = FreeFile

        Set fs = CreateObject("ADODB.Stream")
        fs.Type = 2
        fs.Charset = "utf-8"
        
        'wbName = fs.GetBaseName(ActiveWorkbook.FullName)
        'outFileName = ActiveWorkbook.Path + "\" + wbName + ".txt"
        wbName = Replace(wb.Name, ".xls", "")
        outFileName = wb.Path + "\" + wbName + ".txt"
        'Open outFileName For Output As #nFileNum
        fs.Open

        fs.writetext wbName & Chr(10)
        
        fs.writetext "" & Chr(10)
        fs.writetext "----" & Chr(10)
        fs.writetext "" & Chr(10)

        Dim recInd
        Dim fieldInd
        
        recInd = 0
        For Each myRecord In Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)

            If (recInd > 0) Then
            
                fieldInd = 0
                With myRecord
    
                    For Each myField In Range(.Cells, Cells(.Row, Columns.Count).End(xlToLeft))
                    
                        ' Dot after item number
                        If (fieldInd = 0) Then
                            'Print #nFileNum, myField.Text & "."
                            fs.writetext myField.Text & "." & Chr(10)
                            
                        ' Skip today date
                        ElseIf (fieldInd = 1) Then
                        
                        Else
                            'Print #nFileNum, myField.Text
                            fs.writetext myField.Text & Chr(10)

                        End If
                        
                        ' Extra line after URL
                        If (fieldInd = 9) Then
                            'Print #nFileNum, ""
                            fs.writetext "" & Chr(10)
                        End If
                        
                        fieldInd = fieldInd + 1
                    Next myField
                    
                    'Print #nFileNum, ""
                    'Print #nFileNum, "----"
                    'Print #nFileNum, ""
                    
                    fs.writetext "" & Chr(10)
                    fs.writetext "----" & Chr(10)
                    fs.writetext "" & Chr(10)
    
                End With
                'With myRecord
            
            End If
            ' recInd > 0
            
            recInd = recInd + 1
            
        Next myRecord

        'Close #nFileNum
        fs.SaveToFile outFileName, 2

End Sub




