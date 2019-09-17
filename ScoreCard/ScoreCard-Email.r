library(openxlsx)
library(RDCOMClient)

wb.master <- loadWorkbook(file="V:/Shared Documents/ScoreCard/Master/ScoreCard-Master.xlsx")

readFile.output <- paste("V:/Shared Documents/ScoreCard/Output/ScoreCard-OUTPUT-",as.Date(Sys.Date(),format='%Y%m%d'),".csv",sep="")
data.output <- read.csv(readFile.output, header=TRUE)

readFile.digitalSpend <- paste("V:/Shared Documents/ScoreCard/DigitalAdvertisingSpend/DigitalAdvertisingSpend-Overall.csv",sep="")
data.digitalSpend <- read.csv(readFile.digitalSpend, header=TRUE)

writeData(wb.master,2,data.output)
sheetVisibility(wb.master)[2] <- FALSE
writeData(wb.master,3,data.digitalSpend)
sheetVisibility(wb.master)[3] <- FALSE
writeFile <- paste("V:/Shared Documents/ScoreCard/Report/ScoreCard-",as.Date(Sys.Date(),format='%Y%m%d'),".xlsx",sep="")
saveWorkbook(wb.master,file=writeFile,overwrite=TRUE)

OutApp <- COMCreate("Outlook.Application")

outMail = OutApp$CreateItem(0)
outMail[["SentOnBehalfOfName"]] = "euanm@simpsontravel.com"
outMail[["To"]] = paste("euanm@simpsontravel.com", sep=";", collapse=NULL)
outMail[["Bcc"]] = paste("euanm@simpsontravel.com", sep=";", collapse=NULL)
outMail[["subject"]] = paste("Score Card - ",as.Date(Sys.Date(),format='%Y%m%d'))
outMail[["HTMLbody"]]="<p>Hi all,</p>
<p>Please find attached the latest version of the score card.</p>
<p>Thanks,<br />
Euan Mackenzie<br />
<strong>Digital Enablement Executive</strong></p>
<p>DD&nbsp;&nbsp;&nbsp;&nbsp;020 3319 3770<br />
Tel&nbsp;&nbsp;&nbsp;&nbsp;020 8392 5858<br />
euanm@simpsontravel.com<br />
<a href='https://www.simpsontravel.com'>simpsontravel.com</a><br />
<p>The Grange, Bank of England Sports Ground, Bank Lane, London, SW15 5JT<br />
<img src='https://simpsontravel.blob.core.windows.net/media/26264/emailfooter.png'></p>"
outMail[["attachments"]]$Add(writeFile)
outMail$Send()
