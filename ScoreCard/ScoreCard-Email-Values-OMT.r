library(RDCOMClient)

reportValues.TY <- paste("V:/Shared Documents/ScoreCard/ReportValues/ScoreCard-",format(Sys.Date(),'%Y'),"Departures-",as.Date(Sys.Date(),format='%Y%m%d'),".xlsx",sep="")
reportValues.NY <- paste("V:/Shared Documents/ScoreCard/ReportValues/ScoreCard-",format(Sys.Date()+365,'%Y'),"Departures-",as.Date(Sys.Date(),format='%Y%m%d'),".xlsx",sep="")

OutApp <- COMCreate("Outlook.Application")

outMail = OutApp$CreateItem(0)
outMail[["SentOnBehalfOfName"]] = "data.warehouse@simpsontravel.com"
outMail[["To"]] = paste("OMT@simpsontravel.com", sep=";", collapse=NULL)
outMail[["Bcc"]] = paste("euanm@simpsontravel.com","dave.thomas@clickvisionmedia.co.uk", sep=";", collapse=NULL)
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
outMail[["attachments"]]$Add(reportValues.TY)
outMail[["attachments"]]$Add(reportValues.NY)
outMail$Send()
