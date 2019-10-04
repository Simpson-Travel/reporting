library(openxlsx)
library(RDCOMClient)
DW_RunCheck <- read.xlsx('V:/Shared Documents/DataWarehouse/DW_RunCheck.xlsm',sheet=1)
DW_TablesCheck <- read.xlsx('V:/Shared Documents/DataWarehouse/DW_RunCheck.xlsm',sheet=2)

nrow(DW_TablesCheck)

RunCheck.HasRunToday=0
TableCheck.HasRunToday=0


#If TablesCheck is not blank, and it was refreshed today, TableCheck.HasRunToday=1, else if blank 2 
if(nrow(DW_TablesCheck)>0){
  if(floor(DW_TablesCheck[4])==(as.numeric(as.Date(Sys.Date())-as.Date("1900-01-01"))+2)){
    TableCheck.HasRunToday=1
  }
} else {TableCheck.HasRunToday=2}

if ((DW_RunCheck[1]==0) & (floor(DW_RunCheck[2])==as.numeric(as.Date(Sys.Date())-as.Date("1900-01-01"))+2)) {
  # DW_RunCheck updated, DW was not updated overnight
  
  DW_FailedTables <- DW_TablesCheck[-c(2:4)]
  x=1
  DW_FailedTablesText<-""
  while(x < nrow(DW_FailedTables)){
    DW_FailedTables[x,1]<- paste(DW_FailedTables[x,1],"</li><li>",sep="")
    
    DW_FailedTablesText <- paste(DW_FailedTablesText,DW_FailedTables[x,1],sep="")
    x=x+1
  }
  DW_FailedTables[x,1]<- paste(DW_FailedTables[x,1],"</li>",sep="")
  DW_FailedTablesText <- paste(DW_FailedTablesText,DW_FailedTables[x,1],sep="")
  
  if (TableCheck.HasRunToday==1){
    emailBody <- paste("<p style='font-size:16px'>The Data Warehouse has not completely updated overnight (as of 6:00am).</p><p style='font-size:16px'>The following tables did not update:</p><ul style='font-size:16px'><li>",DW_FailedTablesText,"</ul>", sep="")
  } else {emailBody <- paste("<p style='font-size:16px'>The Data Warehouse has not completely updated overnight (as of 6:00am).</p>", sep="") }
  OutApp <- COMCreate("Outlook.Application")
  outMail = OutApp$CreateItem(0)
  outMail[["SentOnBehalfOfName"]] = "data.warehouse@simpsontravel.com"
  outMail[["To"]] = paste("euanm@simpsontravel.com","simon.rigglesworth@simpsontravel.com","luke.adams@simpsontravel.com", sep=";", collapse=NULL)
  outMail[["subject"]] = paste("Data Warehouse not updated - ",as.Date(Sys.Date(),format='%Y%m%d'))
  outMail[["HTMLbody"]]=emailBody
  outMail$Send()
  
} else if (floor(DW_RunCheck[2])!=as.numeric(as.Date(Sys.Date())-as.Date("1900-01-01"))+2) {
  #Check did not run today
  OutApp <- COMCreate("Outlook.Application")
  outMail = OutApp$CreateItem(0)
  outMail[["SentOnBehalfOfName"]] = "data.warehouse@simpsontravel.com"
  outMail[["To"]] = paste("euanm@simpsontravel.com","simon.rigglesworth@simpsontravel.com","luke.adams@simpsontravel.com", sep=";", collapse=NULL)
  outMail[["subject"]] = paste("Data Warehouse check did not update - ",as.Date(Sys.Date(),format='%Y%m%d'))
  outMail[["HTMLbody"]]="<p style='font-size:16px'>For some reason the Data Warehouse check has failed.<br />Shout at Euan to fix it.</p>"
  outMail$Send()
}


