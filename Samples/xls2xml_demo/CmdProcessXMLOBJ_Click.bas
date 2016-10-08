Private Sub CmdProcessXMLOBJ_Click()


    Dim strXML As String
    If Application.Count(Selection) = 0 Then
        FindUsedRange
    End If
    
    Set ObjXMLDoc = GenerateXMLDOM(Selection, "data")
    
    ObjXMLDoc.Save (filenameinput)
    
    MsgBox ("Completed. XML Written to " & filenameinput)
    Startform.Hide

    

End Sub
