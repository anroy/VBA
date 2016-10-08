Sub Macro1()
'
' Macro1 Macro
'

'
    Workbooks.Open Filename:= _
        "D:\Program Files\iNAGO NetPeople NPMC\Workforms_us_en\Workform_AreaData_us_en\Location_us_en\US-AA.xls"
    ActiveWorkbook.XmlMaps("NPLocationCity_対応付け").Export URL:= _
        "D:\Program Files\iNAGO NetPeople NPMC\XML_us_en\Location_us_en\US-AA.xml"
    ActiveWindow.Close
    Workbooks.Open Filename:= _
        "D:\Program Files\iNAGO NetPeople NPMC\Workforms_us_en\Workform_AreaData_us_en\Location_us_en\US-AE.xls"
    ActiveWorkbook.XmlMaps("NPLocationCity_対応付け").Export URL:= _
        "D:\Program Files\iNAGO NetPeople NPMC\XML_us_en\Location_us_en\US-AE.xml"
    ActiveWindow.Close
    Workbooks.Open Filename:= _
        "D:\Program Files\iNAGO NetPeople NPMC\Workforms_us_en\Workform_AreaData_us_en\Location_us_en\US-AK.xls"
    Range("D8").Select
    ActiveWorkbook.XmlMaps("NPLocationCity_対応付け").Export URL:= _
        "D:\Program Files\iNAGO NetPeople NPMC\XML_us_en\Location_us_en\US-AK.xml"
    ActiveWindow.Close
    Workbooks.Open Filename:= _
        "D:\Program Files\iNAGO NetPeople NPMC\Workforms_us_en\Workform_AreaData_us_en\Location_us_en\US-AL.xls"
    Range("C591").Select
    ActiveWorkbook.XmlMaps("NPLocationCity_対応付け").Export URL:= _
        "D:\Program Files\iNAGO NetPeople NPMC\XML_us_en\Location_us_en\US-AL.xml"
    ActiveWindow.Close
End Sub
