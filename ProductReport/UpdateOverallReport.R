library(openxlsx)
library(RDCOMClient)

wb.master <- loadWorkbook(file="V:/Shared Documents/ProductReport/Master/ProductReport-Master.xlsx")

readFile.output <- paste("V:/Shared Documents/ProductReport/Output/ProductReport-OUTPUT.xlsx",sep="")
data.output <- read.xlsx(readFile.output, sheet=1)

writeData(wb.master,2,data.output)
writeFile.OverallReport <- paste("V:/Shared Documents/ProductReport/Master/ProductReport-Master.xlsx",sep="")
saveWorkbook(wb.master,file=writeFile.OverallReport,overwrite=TRUE)
sheet.names <- getSheetNames(writeFile.OverallReport)

#Mallorca
  wb.Mallorca <- loadWorkbook(file="V:/Shared Documents/ProductReport/Master/ProductReport-Master.xlsx")

  #Committed
  MalCom.Header <- read.xlsx(writeFile.OverallReport, sheet='MalCom', colNames=FALSE, skipEmptyRows=FALSE)
  MalCom.Data <- read.xlsx(writeFile.OverallReport, sheet='MalCom', startRow=4, colNames=FALSE, skipEmptyRows=FALSE)
  writeData(wb.Mallorca,'MalCom',MalCom.Header,startRow=1, colNames=FALSE)
  writeData(wb.Mallorca,'MalCom',MalCom.Data,startRow=4, colNames=FALSE)
  
  #Summary
  MalCom.Header <- read.xlsx(writeFile.OverallReport, sheet='MalComSum', colNames=FALSE, skipEmptyRows=FALSE)
  MalCom.Data <- read.xlsx(writeFile.OverallReport, sheet='MalComSum', startRow=4, colNames=FALSE, skipEmptyRows=FALSE, rows=c(4,5,6))
  writeData(wb.Mallorca,'MalComSum',MalCom.Header,startRow=1, colNames=FALSE)
  writeData(wb.Mallorca,'MalComSum',MalCom.Data,startRow=4, colNames=FALSE)
  
  MalCom.Header <- read.xlsx(writeFile.OverallReport, sheet='MalComSum', startRow=9, colNames=FALSE, skipEmptyRows=FALSE)
  MalCom.Data <- read.xlsx(writeFile.OverallReport, sheet='MalComSum', startRow=12, colNames=FALSE, skipEmptyRows=FALSE)
  writeData(wb.Mallorca,'MalComSum',MalCom.Header,startRow=9, colNames=FALSE)
  writeData(wb.Mallorca,'MalComSum',MalCom.Data,startRow=12, colNames=FALSE)
  
  #On request
  MalCom.Header <- read.xlsx(writeFile.OverallReport, sheet='MalOnReq', colNames=FALSE, skipEmptyRows=FALSE)
  MalCom.Data <- read.xlsx(writeFile.OverallReport, sheet='MalOnReq', startRow=4, colNames=FALSE, skipEmptyRows=FALSE)
  writeData(wb.Mallorca,'MalOnReq',MalCom.Header,startRow=1, colNames=FALSE)
  writeData(wb.Mallorca,'MalOnReq',MalCom.Data,startRow=4, colNames=FALSE)
  
  #Weeks to sell
  MalCom.Header <- read.xlsx(writeFile.OverallReport, sheet='MalWeeksToSell', colNames=FALSE, skipEmptyRows=FALSE)
  MalCom.Data <- read.xlsx(writeFile.OverallReport, sheet='MalWeeksToSell', startRow=3, colNames=FALSE, skipEmptyRows=FALSE)
  writeData(wb.Mallorca,'MalWeeksToSell',MalCom.Header,startRow=1, colNames=FALSE)
  writeData(wb.Mallorca,'MalWeeksToSell',MalCom.Data,startRow=3, colNames=FALSE)
  
  #Remove worksheets
  i = 1 
  while (i<=2){
    removeWorksheet(wb.Mallorca,sheet.names[i])
    i=i+1
  }
  i = 7 
  while (i<=15){
    removeWorksheet(wb.Mallorca,sheet.names[i])
    i=i+1
  }

  writeFile.Mallorca <- paste("V:/Shared Documents/ProductReport/LocationReports/Mallorca-",as.Date(Sys.Date(),format='%Y%m%d'),".xlsx",sep="")
  saveWorkbook(wb.Mallorca,file=writeFile.Mallorca,overwrite=TRUE)