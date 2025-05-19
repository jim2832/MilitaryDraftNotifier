Dim xl
Set xl = CreateObject("Excel.Application")
xl.Visible = False
Set wb = xl.Workbooks.Open("C:\Users\beitun\Desktop\test.xlsm")
xl.Run "CheckSheetAndNotifyLINE"
wb.Close False
xl.Quit
Set wb = Nothing
Set xl = Nothing
