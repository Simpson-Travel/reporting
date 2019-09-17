library(openxlsx)
library(RDCOMClient)

wb.master <- loadWorkbook(file="V:/Shared Documents/ScoreCard/Master/ScoreCard-Master.xlsx")

readFile.output <- paste("V:/Shared Documents/ScoreCard/Output/ScoreCard-OUTPUT-",as.Date(Sys.Date(),format='%Y%m%d'),".csv",sep="")
data.output <- read.csv(readFile.output, header=TRUE)

writeData(wb.master,2,data.output)
writeFile <- paste("C:/Users/euanm/Desktop/ScoreCard-Report/ScoreCard-Reference-",as.Date(Sys.Date(),format='%Y%m%d'),".xlsx",sep="")

sheetVisibility(wb.master)[2] <- FALSE
saveWorkbook(wb.master,file=writeFile,overwrite=TRUE)

excel <- COMCreate("Excel.Application")
excel[["Visible"]] <- TRUE
wb <- excel$Workbooks()$Open(Filename=writeFile)
excel[["DisplayAlerts"]] <- FALSE
wb$Save()
wb$Close(SaveChanges=TRUE)
excel$Quit()

headers <- data.frame(read.xlsx(writeFile, sheet = 1, startRow = 1, colNames = FALSE, skipEmptyRows = FALSE, skipEmptyCols = FALSE))
values <- data.frame(read.xlsx(writeFile, sheet = 1, startRow = 3, colNames = FALSE, skipEmptyRows = FALSE, skipEmptyCols = FALSE))

removeWorksheet(wb.master, 2)
writeData(wb.master,1,headers,startRow=1, colNames=FALSE)
writeData(wb.master,1,values,startRow=3, colNames=FALSE)

writeFile.report <- paste("V:/Shared Documents/ScoreCard/ReportValues/ScoreCard-",as.Date(Sys.Date(),format='%Y%m%d'),".xlsx",sep="")
saveWorkbook(wb.master,file=writeFile.report,overwrite=TRUE)


OutApp <- COMCreate("Outlook.Application")

outMail = OutApp$CreateItem(0)
outMail[["To"]] = paste("euanm@simpsontravel.com","eumackenzie@hotmail.co.uk", sep=";", collapse=NULL)
outMail[["subject"]] = paste("Score Card - ",as.Date(Sys.Date(),format='%Y%m%d'))
outMail[["body"]]="Please find attached the latest version of the score card."
outMail[["attachments"]]$Add(writeFile.report)
outMail$Send()
