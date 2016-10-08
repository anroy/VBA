
Dim arrStates(55)
Dim stateCount

Dim arrCities(55, 10000)
Dim arrLocations(55, 10000)
Dim cityCount(55)


Sub OutputSheet()

    ' Created 2011/10/20 Arka

    Dim bIsAdded
    Dim szCityName, szStateAbbr, szLatitude, szLongitude

    stateCount = 0
    stateIndex = 0
    cityIndex = 0

    Dim rowInd
    rowInd = 2
    Do
        szCityName = Cells(rowInd, 3).Value
        szStateAbbr = Cells(rowInd, 8).Value
        szLatitude = Cells(rowInd, 15).Value
        szLongitude = Cells(rowInd, 16).Value
        
        stateIndex = getStateIndex(szStateAbbr)

        bIsAdded = cityAdded(stateIndex, szCityName)
        If Not bIsAdded Then
            arrCities(stateIndex, cityCount(stateIndex)) = szCityName
            arrLocations(stateIndex, cityCount(stateIndex)) = CStr(szLatitude) + "," + CStr(szLongitude)
            cityCount(stateIndex) = cityCount(stateIndex) + 1
        End If
        
        rowInd = rowInd + 1
    Loop Until IsEmpty(Cells(rowInd, 1))


    ' write out files
    
    Dim outputLine
    Dim CriteriaName, CriteriaValue, Normal, ParentNormal, Location, Proximity, BoundingRegion, ValueLabel, ProximityLabel, SelectionLabel, Synonyms
    CriteriaName = "City"
    
    Proximity = "$CITY_PROX"
    BoundingRegion = ""
    ValueLabel = ""
    ProximityLabel = "$IN_DESC"

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    For stateIndex = 0 To stateCount - 1
    
        ParentNormal = "US-" + arrStates(stateIndex)

        Set fs = CreateObject("ADODB.Stream")
        fs.Type = 2
        fs.Charset = "utf-8"
        wbName = Replace(wb.Name, ".xls", "")
        outFileName = wb.Path + "\CityFiles\" + ParentNormal + ".csv"
        fs.Open
        fs.writetext "CriteriaName,CriteriaValue,Normal,ParentNormal,Location,Proximity,BoundingRegion,ValueLabel,ProximityLabel,SelectionLabel,Synonyms" & Chr(10)


        For cityIndex = 0 To cityCount(stateIndex) - 1
        
            CriteriaValue = arrCities(stateIndex, cityIndex)
            Normal = CriteriaValue
            
            Location = """" + arrLocations(stateIndex, cityIndex) + """"

            SelectionLabel = CriteriaValue
            Synonyms = CriteriaValue
    
            outputLine = CriteriaName + "," + CriteriaValue + "," + Normal + "," + ParentNormal + "," + Location + "," + Proximity + "," + BoundingRegion + "," + ValueLabel + "," + ProximityLabel + "," + SelectionLabel + "," + Synonyms
            fs.writetext outputLine & Chr(10)
    
        Next cityIndex
        
        fs.SaveToFile outFileName, 2
        fs.Close
        
    Next stateIndex


End Sub


Function getStateIndex(szStateAbbr) As Integer

    bStateAdded = False

    For i = 0 To stateCount - 1

        If arrStates(i) = szStateAbbr Then
            bStateAdded = True
            getStateIndex = i
            Exit For
        End If
    Next i
    
    If Not bStateAdded Then
        arrStates(stateCount) = szStateAbbr
        getStateIndex = stateCount
        stateCount = stateCount + 1
    End If

End Function


Function cityAdded(stateIndex, CityName) As Boolean

    cityAdded = False

    For i = 0 To cityCount(stateIndex) - 1

        If arrCities(stateIndex, i) = CityName Then
            cityAdded = True
            Exit For
        End If

    Next i

End Function

