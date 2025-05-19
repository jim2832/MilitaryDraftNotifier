Dim xl
Set xl = CreateObject("Excel.Application")
xl.Visible = False
Set wb = xl.Workbooks.Open("C:\Users\beitun\Desktop\out.xlsm")
xl.Run "CheckColumnPairsAndNotifyLINE"
wb.Close False
xl.Quit
Set xl = Nothing
