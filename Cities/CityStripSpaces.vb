Sub CityStripSpaces()

    ' Created 2011/10/25 Arka

    Dim rowInd
    rowInd = 2
    Do
        OldNormal = Cells(rowInd, 3).Value
        NewNormal = Replace(OldNormal, " ", "")
        Cells(rowInd, 3).Value = NewNormal
        
        rowInd = rowInd + 1
    Loop Until IsEmpty(Cells(rowInd, 1))

End Sub

