ExcelFilePath = "V:\Shared Documents\Aviation\Master\Aviation-FlightReport.xlsm"
MacroPath = "Module1.FlightReport"
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible = False
Set wb = ExcelApp.Workbooks.Open(ExcelFilePath)
ExcelApp.Run MacroPath
wb.Save
wb.Close
ExcelApp.Quit
