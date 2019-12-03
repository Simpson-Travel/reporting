library(openxlsx)
library(RDCOMClient)


outputFile.date <- "V:/Shared Documents/BedBalance/Output/BedBalance-Date.csv"
outputTable.date <- read.csv(outputFile.date,header=TRUE)

if (as.Date(outputTable.date$Date[1])==Sys.Date()){

  outputFile.flights <- "V:/Shared Documents/BedBalance/Output/BedBalance-Flights.csv"
  outputFile.beds <- "V:/Shared Documents/BedBalance/Output/BedBalance-Beds.csv"
  
  outputTable.flights <- read.csv(outputFile.flights,header=TRUE)
  outputTable.beds <- read.csv(outputFile.beds,header=TRUE)
  
  writeFile <- "V:/Shared Documents/BedBalance/Master/BedBalance-Master.xlsx"
  
  wb.BedBalanceMaster <- loadWorkbook(writeFile)
  deleteData(wb.BedBalanceMaster,"Flights",rows=1:100000,cols=1:200,gridExpand=TRUE)
  deleteData(wb.BedBalanceMaster,"Beds",rows=1:100000,cols=1:200,gridExpand=TRUE)
  writeData(wb.BedBalanceMaster,"Flights",outputTable.flights,rowNames=FALSE)
  writeData(wb.BedBalanceMaster,"Beds",outputTable.beds,rowNames=FALSE)
  saveWorkbook(wb.BedBalanceMaster,writeFile,overwrite=TRUE)
  
  excel <- COMCreate("Excel.Application")
  excel[["Visible"]] <- TRUE
  wb <- excel$Workbooks()$Open(Filename=writeFile)
  excel[["DisplayAlerts"]] <- FALSE
  wb$CalculateFull
  wb$Save()
  wb$Close(SaveChanges=TRUE)
  excel$Quit()
  
  wb.BedBalanceReport <- loadWorkbook(writeFile)
  Table.BedBalance <- readWorkbook(wb.BedBalanceReport,sheet="BedBalance",colNames=TRUE,skipEmptyRows=FALSE, startRow=2)
  Table.BedBalance <- Table.BedBalance[-c(1:2)]
  writeData(wb.BedBalanceReport,"BedBalanceReport",Table.BedBalance,colNames=FALSE,rowNames=FALSE,startRow=2)
  removeWorksheet(wb.BedBalanceReport,"Flights")
  removeWorksheet(wb.BedBalanceReport,"Beds")
  removeWorksheet(wb.BedBalanceReport,"Master")
  removeWorksheet(wb.BedBalanceReport,"BedBalance")
  writeFile.report <- paste("V:/Shared Documents/BedBalance/Report/BedBalance-Report-", format(Sys.Date(),"%Y%m%d"),".xlsx",sep="")
  saveWorkbook(wb.BedBalanceReport,writeFile.report,overwrite=TRUE)
  
   OutApp <- COMCreate("Outlook.Application")
    
   outMail = OutApp$CreateItem(0)
   outMail[["SentOnBehalfOfName"]] = "data.warehouse@simpsontravel.com"
   outMail[["To"]] = paste("daniel@simpsontravel.com","phil@simpsontravel.com","michelle@simpsontravel.com", sep=";", collapse=NULL)
   outMail[["Bcc"]] = paste("euanm@simpsontravel.com", sep=";", collapse=NULL)
   outMail[["subject"]] = paste("Bed Balance - ",as.Date(Sys.Date(),format='%Y-%m-%d'))
   outMail[["HTMLbody"]]="<p>Hi all,</p>
     <p>Please find attached today's Bed Balance report.</p>
     <p>Thanks,<br />
     Euan Mackenzie<br />
     <strong>Digital Enablement Executive</strong></p>
     <p>DD&nbsp;&nbsp;&nbsp;&nbsp;020 3319 3770<br />
     Tel&nbsp;&nbsp;&nbsp;&nbsp;020 8392 5858<br />
     euanm@simpsontravel.com<br />
     <a href='https://www.simpsontravel.com'>simpsontravel.com</a><br />
     <p>The Grange, Bank of England Sports Ground, Bank Lane, London, SW15 5JT<br />
     <img src='https://simpsontravel.blob.core.windows.net/media/26264/emailfooter.png'></p>"
    outMail[["attachments"]]$Add(writeFile.report)
    outMail$Send()

}