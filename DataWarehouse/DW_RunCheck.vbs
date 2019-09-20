ExcelFilePath = "V:\Shared Documents\DataWarehouse\DW_RunCheck.xlsm"
MacroPath = "Module1.DW_RunCheck"
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible = False
Set wb = ExcelApp.Workbooks.Open(ExcelFilePath)
ExcelApp.Run MacroPath
wb.Save
wb.Close
ExcelApp.Quit
