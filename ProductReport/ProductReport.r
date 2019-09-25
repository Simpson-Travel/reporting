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
  Header <- read.xlsx(writeFile.OverallReport, sheet='MalCom', colNames=FALSE, skipEmptyRows=FALSE)
  Data <- read.xlsx(writeFile.OverallReport, sheet='MalCom', startRow=4, colNames=FALSE, skipEmptyRows=FALSE)
  writeData(wb.Mallorca,'MalCom',Header,startRow=1, colNames=FALSE)
  writeData(wb.Mallorca,'MalCom',Data,startRow=4, colNames=FALSE)
  setColWidths(wb.Mallorca,'MalCom',cols=1:1,widths='auto',hidden=TRUE)

  #Summary
  Header <- read.xlsx(writeFile.OverallReport, sheet='MalComSum', colNames=FALSE, skipEmptyRows=FALSE)
  Data <- read.xlsx(writeFile.OverallReport, sheet='MalComSum', startRow=4, colNames=FALSE, skipEmptyRows=FALSE, rows=c(4,5,6))
  writeData(wb.Mallorca,'MalComSum',Header,startRow=1, colNames=FALSE)
  writeData(wb.Mallorca,'MalComSum',Data,startRow=4, colNames=FALSE)

  Header <- read.xlsx(writeFile.OverallReport, sheet='MalComSum', startRow=9, colNames=FALSE, skipEmptyRows=FALSE)
  Data <- read.xlsx(writeFile.OverallReport, sheet='MalComSum', startRow=12, colNames=FALSE, skipEmptyRows=FALSE)
  writeData(wb.Mallorca,'MalComSum',Header,startRow=9, colNames=FALSE)
  writeData(wb.Mallorca,'MalComSum',Data,startRow=12, colNames=FALSE)

  #On request
  Header <- read.xlsx(writeFile.OverallReport, sheet='MalOnReq', colNames=FALSE, skipEmptyRows=FALSE)
  Data <- read.xlsx(writeFile.OverallReport, sheet='MalOnReq', startRow=4, colNames=FALSE, skipEmptyRows=FALSE)
  writeData(wb.Mallorca,'MalOnReq',Header,startRow=1, colNames=FALSE)
  writeData(wb.Mallorca,'MalOnReq',Data,startRow=4, colNames=FALSE)
  setColWidths(wb.Mallorca,'MalOnReq',cols=1:1,widths='auto',hidden=TRUE)

  #Weeks to sell
  Header <- read.xlsx(writeFile.OverallReport, sheet='MalWeeksToSell', colNames=FALSE, skipEmptyRows=FALSE)
  Data <- read.xlsx(writeFile.OverallReport, sheet='MalWeeksToSell', startRow=3, colNames=FALSE, skipEmptyRows=FALSE)
  writeData(wb.Mallorca,'MalWeeksToSell',Header,startRow=1, colNames=FALSE)
  writeData(wb.Mallorca,'MalWeeksToSell',Data,startRow=3, colNames=FALSE)
  setColWidths(wb.Mallorca,'MalWeeksToSell',cols=1:1,widths='auto',hidden=TRUE)

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

#Turkey
  wb.Turkey <- loadWorkbook(file="V:/Shared Documents/ProductReport/Master/ProductReport-Master.xlsx")

  #Committed
    #Islamlar
    Header <- read.xlsx(writeFile.OverallReport, sheet='IslamlarCom', colNames=FALSE, skipEmptyRows=FALSE)
    Data <- read.xlsx(writeFile.OverallReport, sheet='IslamlarCom', startRow=4, colNames=FALSE, skipEmptyRows=FALSE)
    writeData(wb.Turkey,'IslamlarCom',Header,startRow=1, colNames=FALSE)
    writeData(wb.Turkey,'IslamlarCom',Data,startRow=4, colNames=FALSE)
    setColWidths(wb.Turkey,'IslamlarCom',cols=1:1,widths='auto',hidden=TRUE)

    #Kalkan
    Header <- read.xlsx(writeFile.OverallReport, sheet='KalkanCom', colNames=FALSE, skipEmptyRows=FALSE)
    Data <- read.xlsx(writeFile.OverallReport, sheet='KalkanCom', startRow=4, colNames=FALSE, skipEmptyRows=FALSE)
    writeData(wb.Turkey,'KalkanCom',Header,startRow=1, colNames=FALSE)
    writeData(wb.Turkey,'KalkanCom',Data,startRow=4, colNames=FALSE)
    setColWidths(wb.Turkey,'KalkanCom',cols=1:1,widths='auto',hidden=TRUE)

    #Akyaka
    Header <- read.xlsx(writeFile.OverallReport, sheet='AkyakaCom', colNames=FALSE, skipEmptyRows=FALSE)
    Data <- read.xlsx(writeFile.OverallReport, sheet='AkyakaCom', startRow=4, colNames=FALSE, skipEmptyRows=FALSE)
    writeData(wb.Turkey,'AkyakaCom',Header,startRow=1, colNames=FALSE)
    writeData(wb.Turkey,'AkyakaCom',Data,startRow=4, colNames=FALSE)
    setColWidths(wb.Turkey,'AkyakaCom',cols=1:1,widths='auto',hidden=TRUE)

    #Bozburun
    Header <- read.xlsx(writeFile.OverallReport, sheet='BozburunCom', colNames=FALSE, skipEmptyRows=FALSE)
    Data <- read.xlsx(writeFile.OverallReport, sheet='BozburunCom', startRow=4, colNames=FALSE, skipEmptyRows=FALSE)
    writeData(wb.Turkey,'BozburunCom',Header,startRow=1, colNames=FALSE)
    writeData(wb.Turkey,'BozburunCom',Data,startRow=4, colNames=FALSE)
    setColWidths(wb.Turkey,'BozburunCom',cols=1:1,widths='auto',hidden=TRUE)

    #Dalyan
    Header <- read.xlsx(writeFile.OverallReport, sheet='DalyanCom', colNames=FALSE, skipEmptyRows=FALSE)
    Data <- read.xlsx(writeFile.OverallReport, sheet='DalyanCom', startRow=4, colNames=FALSE, skipEmptyRows=FALSE)
    writeData(wb.Turkey,'DalyanCom',Header,startRow=1, colNames=FALSE)
    writeData(wb.Turkey,'DalyanCom',Data,startRow=4, colNames=FALSE)
    setColWidths(wb.Turkey,'DalyanCom',cols=1:1,widths='auto',hidden=TRUE)

    #Kas
    Header <- read.xlsx(writeFile.OverallReport, sheet='KasCom', colNames=FALSE, skipEmptyRows=FALSE)
    Data <- read.xlsx(writeFile.OverallReport, sheet='KasCom', startRow=4, colNames=FALSE, skipEmptyRows=FALSE)
    writeData(wb.Turkey,'KasCom',Header,startRow=1, colNames=FALSE)
    writeData(wb.Turkey,'KasCom',Data,startRow=4, colNames=FALSE)
    setColWidths(wb.Turkey,'KasCom',cols=1:1,widths='auto',hidden=TRUE)

  #Summary
  Header <- read.xlsx(writeFile.OverallReport, sheet='TurkeyComSum', colNames=FALSE, skipEmptyRows=FALSE)
  Data <- read.xlsx(writeFile.OverallReport, sheet='TurkeyComSum', startRow=4, colNames=FALSE, skipEmptyRows=FALSE, rows=c(4,5,6,7,8,9,10,11))
  writeData(wb.Turkey,'TurkeyComSum',Header,startRow=1, colNames=FALSE)
  writeData(wb.Turkey,'TurkeyComSum',Data,startRow=4, colNames=FALSE)

  Header <- read.xlsx(writeFile.OverallReport, sheet='TurkeyComSum', startRow=14, colNames=FALSE, skipEmptyRows=FALSE)
  Data <- read.xlsx(writeFile.OverallReport, sheet='TurkeyComSum', startRow=17, colNames=FALSE, skipEmptyRows=FALSE)
  writeData(wb.Turkey,'TurkeyComSum',Header,startRow=14, colNames=FALSE)
  writeData(wb.Turkey,'TurkeyComSum',Data,startRow=17, colNames=FALSE)

  #On request
  Header <- read.xlsx(writeFile.OverallReport, sheet='TurkeyOnReq', colNames=FALSE, skipEmptyRows=FALSE)
  Data <- read.xlsx(writeFile.OverallReport, sheet='TurkeyOnReq', startRow=4, colNames=FALSE, skipEmptyRows=FALSE)
  writeData(wb.Turkey,'TurkeyOnReq',Header,startRow=1, colNames=FALSE)
  writeData(wb.Turkey,'TurkeyOnReq',Data,startRow=4, colNames=FALSE)
  setColWidths(wb.Turkey,'TurkeyOnReq',cols=1:1,widths='auto',hidden=TRUE)

  #Weeks to sell
  Header <- read.xlsx(writeFile.OverallReport, sheet='TurkeyWeeksToSell', colNames=FALSE, skipEmptyRows=FALSE)
  Data <- read.xlsx(writeFile.OverallReport, sheet='TurkeyWeeksToSell', startRow=3, colNames=FALSE, skipEmptyRows=FALSE)
  writeData(wb.Turkey,'TurkeyWeeksToSell',Header,startRow=1, colNames=FALSE)
  writeData(wb.Turkey,'TurkeyWeeksToSell',Data,startRow=3, colNames=FALSE)
  setColWidths(wb.Turkey,'TurkeyWeeksToSell',cols=1:1,widths='auto',hidden=TRUE)

  #Remove worksheets
  i = 1
  while (i<=6){
    removeWorksheet(wb.Turkey,sheet.names[i])
    i=i+1
  }
  #i = 7
  #while (i<=15){
  #  removeWorksheet(wb.Turkey,sheet.names[i])
  #  i=i+1
  #}

  writeFile.Turkey <- paste("V:/Shared Documents/ProductReport/LocationReports/Turkey-",as.Date(Sys.Date(),format='%Y%m%d'),".xlsx",sep="")
  saveWorkbook(wb.Turkey,file=writeFile.Turkey,overwrite=TRUE)
