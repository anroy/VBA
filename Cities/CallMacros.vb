Sub AAAAA()
'
' AAAAA Macro
'

'
    Workbooks.Open Filename:="C:\Users\arka.roy\Desktop\yama\US-DC.xls"
    Application.Run "Scripts.xls!CityStripSpaces"
    ActiveWorkbook.Save
    ActiveWorkbook.Close
	
    Workbooks.Open Filename:="C:\Users\arka.roy\Desktop\yama\US-GU.xls"
    Application.Run "Scripts.xls!CityStripSpaces"
    ActiveWorkbook.Save
    ActiveWorkbook.Close
	
    Workbooks.Open Filename:="C:\Users\arka.roy\Desktop\yama\US-HI.xls"
    ActiveWindow.LargeScroll ToRight:=-1
    Application.Run "Scripts.xls!CityStripSpaces"
    ActiveWorkbook.Save
    ActiveWorkbook.Close
	

End Sub
