library(reshape2)
library(openxlsx)

beds <- read.csv(file="V:/Shared Documents/BedBalance/Output/BedBalance-Beds.csv", header=TRUE)

colnames(beds)[colnames(beds)=="AdjustedWeekStartDate"] <- "Date"
colnames(beds)[colnames(beds)=="Accom.Location.CountrySourceID"] <- "CountryID"
colnames(beds)[colnames(beds)=="Accom.Location.CountryName"] <- "CountryName"
colnames(beds)[colnames(beds)=="Accom.Location.RegionSourceID"] <- "RegionID"
colnames(beds)[colnames(beds)=="Accom.Location.RegionName"] <- "RegionName"
colnames(beds)[colnames(beds)=="Accom.Occ.maximumOccupancy"] <- "MaxOcc"
colnames(beds)[colnames(beds)=="Accom.ParentAccom.AccommodationType"] <- "AccomType"

availability <- subset(beds,as.Date(beds$Date, format="%d/%m/%Y")>=as.Date("01/11/2019",format="%d/%m/%Y") & as.Date(beds$Date, format="%d/%m/%Y")<as.Date("01/11/2020",format="%d/%m/%Y") & MaxOcc>0 & AvailableFlag.Committed>0)
availability$Date <- as.Date(availability$Date,"%d/%m/%Y")
availability <- availability[order(availability$Date),]

# COUNTRY AND REGION REPORT
availability.Region <- aggregate(AvailableFlag.Committed ~ Date+CountryID+CountryName+RegionID+RegionName+MaxOcc, availability, sum)
availability.Region <- recast (availability.Region,CountryID+CountryName+RegionID+RegionName+MaxOcc ~ Date,id.var=c("CountryID","CountryName","RegionID","RegionName","MaxOcc","Date"))
availability.Region[is.na(availability.Region)] <- 0
availability.Region <- subset(availability.Region,select=-c(CountryID,RegionID))

#writeFile.Region <- "C:/Users/euanm/Desktop/BedsByOccupancy-Region.csv"
#write.csv(availability.Region,writeFile.Region,row.names=FALSE)

# COUNTRY REPORT
availability.Country <- aggregate(AvailableFlag.Committed ~ Date+CountryID+CountryName+MaxOcc, availability, sum)
availability.Country <- recast (availability.Country,CountryID+CountryName+MaxOcc ~ Date,id.var=c("CountryID","CountryName","MaxOcc","Date"))
availability.Country[is.na(availability.Country)] <- 0
availability.Country <- subset(availability.Country,select=-c(CountryID))

# Hotel REPORT
availability.Hotel <- subset(aggregate(AvailableFlag.Committed ~ Date+CountryID+CountryName+RegionID+RegionName+AccomType+MaxOcc, availability, sum),AccomType=="Hotel")
availability.Hotel <- recast (availability.Hotel,CountryID+CountryName+RegionID+RegionName+AccomType+MaxOcc ~ Date,id.var=c("CountryID","CountryName","RegionID","RegionName","AccomType","MaxOcc","Date"))
availability.Hotel[is.na(availability.Hotel)] <- 0
availability.Hotel <- subset(availability.Hotel,select=-c(CountryID,RegionID))

#writeFile.Country <- "C:/Users/euanm/Desktop/BedsByOccupancy-Country.csv"
#write.csv(availability.Country,writeFile.Country,row.names=FALSE)

BedsByOccupancy.Read <- "V:/Shared Documents/BedBalance/ByOccupancy/BedsByOccupancy.xlsx"
BedsByOccupancy.Write <- paste("V:/Shared Documents/BedBalance/ByOccupancy/Report/BedsByOccupancy-",format(Sys.Date(),"%Y%m%d"),".xlsx",sep="")

wb.BedsByOccupancyMaster <- loadWorkbook(BedsByOccupancy.Read)
hs1 <- createStyle(textDecoration="Bold")
deleteData(wb.BedsByOccupancyMaster,"Country",cols=1:1000,rows=1:1000,gridExpand=TRUE)
writeData(wb.BedsByOccupancyMaster,"Country",availability.Country,rowNames=FALSE,headerStyle=hs1)
setColWidths(wb.BedsByOccupancyMaster,"Country",cols=1:1000,widths="auto")

deleteData(wb.BedsByOccupancyMaster,"Region",cols=1:1000,rows=1:1000,gridExpand=TRUE)
writeData(wb.BedsByOccupancyMaster,"Region",availability.Region,rowNames=FALSE,headerStyle=hs1)
setColWidths(wb.BedsByOccupancyMaster,"Region",cols=1:1000,widths="auto")

deleteData(wb.BedsByOccupancyMaster,"Hotels",cols=1:1000,rows=1:1000,gridExpand=TRUE)
writeData(wb.BedsByOccupancyMaster,"Hotels",availability.Hotel,rowNames=FALSE,headerStyle=hs1)
setColWidths(wb.BedsByOccupancyMaster,"Hotels",cols=1:1000,widths="auto")

saveWorkbook(wb.BedsByOccupancyMaster,BedsByOccupancy.Write,overwrite=TRUE)
