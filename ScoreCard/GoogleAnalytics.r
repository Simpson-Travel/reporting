library(RGoogleAnalytics)
library(RGA)
library(lubridate)

# Map https://simpsontravelco.sharepoint.com/sites/Digital as a network drive to V:

# GA authorisation
load("V:/Shared Documents/ScoreCard/GoogleAnalytics/RScript/token_file")
ValidateToken(token)

# MCF authorisation
authorize()

# Dates
monday.TW <- as.Date(cut(as.Date(Sys.Date()),"week"))-7
sunday.TW <- as.Date(cut(as.Date(Sys.Date()),"week"))-1
monday.LW <- as.Date(cut(as.Date(Sys.Date()),"week"))-14
sunday.LW <- as.Date(cut(as.Date(Sys.Date()),"week"))-8
TM <- as.numeric(format(monday.TW,'%m'))
TY <- as.numeric(format(monday.TW,'%Y'))
LY <- as.numeric(format(monday.TW,'%Y'))-1
isoWeek.LY <- as.numeric(paste(LY,isoweek(monday.TW),sep=""))
startDate.TMTY <- as.Date(paste(TY,TM,1,sep="-"))
startDate.TMLY <- as.Date(paste(LY,TM,1,sep="-"))
endDate.TMTY <- if (TM != as.numeric(format(Sys.Date(),'%m'))) {as.Date(format(Sys.Date(),'%Y-%m-01'))-1} else {as.Date(format(Sys.Date(),'%Y-%m-%d'))-1}
endDate.TMLY <- as.Date(paste(as.numeric(format(endDate.TMTY,'%Y'))-1,format(endDate.TMTY,'%m'),format(endDate.TMTY,'%d%'),sep="-"))
FY <- if (TM >= 11){as.numeric(format(monday.TW,'%Y'))+1} else {as.numeric(format(monday.TW,'%Y'))}
startDate.FYTY <- as.Date(paste(FY-1,11,1,sep="-"))
startDate.FYLY <- as.Date(paste(FY-2,11,1,sep="-"))
endDate.FYTY <- as.Date(format(Sys.Date(),'%Y-%m-%d'))-1
endDate.FYLY <- as.Date(paste(LY,format(Sys.Date(),'-%m-%d'),sep=""))-1

# Date Text
TW.startDate <- format(monday.TW,'%Y-%m-%d')
TW.endDate <- format(sunday.TW,'%Y-%m-%d')
LW.startDate <- format(monday.LW,'%Y-%m-%d')
LW.endDate <- format(sunday.LW,'%Y-%m-%d')
TMTY.startDate <- format(startDate.TMTY,'%Y-%m-%d')
TMTY.endDate <- format(endDate.TMTY,'%Y-%m-%d')
TMLY.startDate <- format(startDate.TMLY,'%Y-%m-%d')
TMLY.endDate <- format(endDate.TMLY,'%Y-%m-%d')
FYTY.startDate <- format(startDate.FYTY,'%Y-%m-%d')
FYTY.endDate <- format(endDate.FYTY,'%Y-%m-%d')
FYLY.startDate <- format(startDate.FYLY,'%Y-%m-%d')
FYLY.endDate <- format(endDate.FYLY,'%Y-%m-%d')

# 1_1_1 - GA - Users - TW
Users.TW.query.init <- Init(start.date = TW.startDate,
                            end.date = TW.endDate,
                            dimensions = c("ga:isoYearIsoWeek"),
                            metrics = c("ga:users"),
                            max.results=10000,
                            table.id = "ga:149139688")
Users.TW.query <- QueryBuilder(Users.TW.query.init)
Users.TW.data <- GetReportData(Users.TW.query, token, paginate_query = TRUE)

# 1_1_2 - GA - Users - LW
Users.LW.query.init <- Init(start.date = LW.startDate,
                            end.date = LW.endDate,
                            dimensions = c("ga:isoYearIsoWeek"),
                            metrics = c("ga:users"),
                            max.results=10000,
                            table.id = "ga:149139688")
Users.LW.query <- QueryBuilder(Users.LW.query.init)
Users.LW.data <- GetReportData(Users.LW.query, token, paginate_query = TRUE)

# 1_1_3 - GA - Users - TWLY
Users.TWLY.query.init <- Init(start.date = "2017-04-01",
                              end.date = FYTY.endDate,
                              dimensions = c("ga:isoYearIsoWeek"),
                              metrics = c("ga:users"),
                              filters = paste("ga:users>0;ga:isoYearIsoWeek==",isoWeek.LY,sep=""),
                              max.results=10000,
                              table.id = "ga:149139688")
Users.TWLY.query <- QueryBuilder(Users.TWLY.query.init)
Users.TWLY.data <- GetReportData(Users.TWLY.query, token, paginate_query = TRUE)

# 1_1 - GA - Users - TW
Users.Week.data <- rbind(Users.TW.data,Users.LW.data,Users.TWLY.data)
write.csv(Users.Week.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/1_1-Users-TW.csv",row.names=FALSE,append=FALSE,na="")

# 1_2 - GA - Users - TM
Users.TM.query.init <- Init(start.date = TMTY.startDate,
                            end.date = TMTY.endDate,
                            dimensions = c("ga:yearMonth"),
                            metrics = c("ga:users"),
                            max.results=10000,
                            table.id = "ga:149139688")
Users.TM.query <- QueryBuilder(Users.TM.query.init)
Users.TM.data <- GetReportData(Users.TM.query, token, paginate_query = TRUE)
write.csv(Users.TM.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/1_2-Users-TM.csv",row.names=FALSE,append=FALSE,na="")

# 1_3 - GA - Users - TM LY
Users.TMLY.query.init <- Init(start.date = TMLY.startDate,
                              end.date = TMLY.endDate,
                              dimensions = c("ga:yearMonth"),
                              metrics = c("ga:users"),
                              max.results=10000,
                              table.id = "ga:149139688")
Users.TMLY.query <- QueryBuilder(Users.TMLY.query.init)
Users.TMLY.data <- GetReportData(Users.TMLY.query, token, paginate_query = TRUE)
write.csv(Users.TMLY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/1_3-Users-TMLY.csv",row.names=FALSE,append=FALSE,na="")

# 1_4 - GA - Users - TY
Users.TY.query.init <- Init(start.date = FYTY.startDate,
                            end.date = FYTY.endDate,
                            dimensions = c(""),
                            metrics = c("ga:users"),
                            max.results=10000,
                            table.id = "ga:149139688")
Users.TY.query <- QueryBuilder(Users.TY.query.init)
Users.TY.data <- GetReportData(Users.TY.query, token, paginate_query = TRUE)
write.csv(Users.TY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/1_4-Users-TY.csv",row.names=FALSE,append=FALSE,na="")

# 1_5 - GA - Users - LY
Users.LY.query.init <- Init(start.date = FYLY.startDate,
                            end.date = FYLY.endDate,
                            dimensions = c(""),
                            metrics = c("ga:users"),
                            max.results=10000,
                            table.id = "ga:149139688")
Users.LY.query <- QueryBuilder(Users.LY.query.init)
Users.LY.data <- GetReportData(Users.LY.query, token, paginate_query = TRUE)
write.csv(Users.LY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/1_5-Users-LY.csv",row.names=FALSE,append=FALSE,na="")

#2_1_1 - GA - Paid Users - TW
PaidUsers.TW.query.init <- Init(start.date = TW.startDate,
                                end.date = TW.endDate,
                                dimensions = c("ga:isoYearIsoWeek"),
                                metrics = c("ga:users"),
                                filters = "ga:channelGrouping==Paid Search,ga:channelGrouping==Display,ga:source==criteo,ga:medium==paid-social,ga:medium==paid",
                                max.results=10000,
                                table.id = "ga:149139688")
PaidUsers.TW.query <- QueryBuilder(PaidUsers.TW.query.init)
PaidUsers.TW.data <- GetReportData(PaidUsers.TW.query, token, paginate_query = TRUE)

# 2_1_2 - GA - Paid Users - LW
PaidUsers.LW.query.init <- Init(start.date = LW.startDate,
                                end.date = LW.endDate,
                                dimensions = c("ga:isoYearIsoWeek"),
                                metrics = c("ga:users"),
                                filters = "ga:channelGrouping==Paid Search,ga:channelGrouping==Display,ga:source==criteo,ga:medium==paid-social,ga:medium==paid",
                                max.results=10000,
                                table.id = "ga:149139688")
PaidUsers.LW.query <- QueryBuilder(PaidUsers.LW.query.init)
PaidUsers.LW.data <- GetReportData(PaidUsers.LW.query, token, paginate_query = TRUE)

# 2_1_3 - GA - Paid Users - TWLY
PaidUsers.TWLY.query.init <- Init(start.date = "2017-04-01",
                                  end.date = FYTY.endDate,
                                  dimensions = c("ga:isoYearIsoWeek"),
                                  metrics = c("ga:users"),
                                  filters = paste("ga:users>0;ga:channelGrouping==Paid Search,ga:channelGrouping==Display,ga:source==criteo,ga:medium==paid-social,ga:medium==paid;ga:isoYearIsoWeek==",isoWeek.LY,sep=""),
                                  max.results=10000,
                                  table.id = "ga:149139688")
PaidUsers.TWLY.query <- QueryBuilder(PaidUsers.TWLY.query.init)
PaidUsers.TWLY.data <- GetReportData(PaidUsers.TWLY.query, token, paginate_query = TRUE)

# 2_1 - GA - Paid Users - TW
PaidUsers.Week.data <- rbind(PaidUsers.TW.data,PaidUsers.LW.data,PaidUsers.TWLY.data)
write.csv(PaidUsers.Week.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/2_1-PaidUsers-TW.csv",row.names=FALSE,append=FALSE,na="")

# 2_2 - GA - Paid Users - TM
PaidUsers.TM.query.init <- Init(start.date = TMTY.startDate,
                                end.date = TMTY.endDate,
                                dimensions = c("ga:yearMonth"),
                                metrics = c("ga:users"),
                                filters = "ga:channelGrouping==Paid Search,ga:channelGrouping==Display,ga:source==criteo,ga:medium==paid-social,ga:medium==paid",
                                max.results=10000,
                                table.id = "ga:149139688")
PaidUsers.TM.query <- QueryBuilder(PaidUsers.TM.query.init)
PaidUsers.TM.data <- GetReportData(PaidUsers.TM.query, token, paginate_query = TRUE)
write.csv(PaidUsers.TM.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/2_2-PaidUsers-TM.csv",row.names=FALSE,append=FALSE,na="")

# 2_3 - GA - Paid Users - TM LY
PaidUsers.TMLY.query.init <- Init(start.date = TMLY.startDate,
                                  end.date = TMLY.endDate,
                                  dimensions = c("ga:yearMonth"),
                                  metrics = c("ga:users"),
                                  filters = "ga:channelGrouping==Paid Search,ga:channelGrouping==Display,ga:source==criteo,ga:medium==paid-social,ga:medium==paid",
                                  max.results=10000,
                                  table.id = "ga:149139688")
PaidUsers.TMLY.query <- QueryBuilder(PaidUsers.TMLY.query.init)
PaidUsers.TMLY.data <- GetReportData(PaidUsers.TMLY.query, token, paginate_query = TRUE)
write.csv(PaidUsers.TMLY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/2_3-PaidUsers-TMLY.csv",row.names=FALSE,append=FALSE,na="")

# 2_4 - GA - Paid Users - TY
PaidUsers.TY.query.init <- Init(start.date = FYTY.startDate,
                                end.date = FYTY.endDate,
                                dimensions = c(""),
                                metrics = c("ga:users"),
                                filters = "ga:channelGrouping==Paid Search,ga:channelGrouping==Display,ga:source==criteo,ga:medium==paid-social,ga:medium==paid",
                                max.results=10000,
                                table.id = "ga:149139688")
PaidUsers.TY.query <- QueryBuilder(PaidUsers.TY.query.init)
PaidUsers.TY.data <- GetReportData(PaidUsers.TY.query, token, paginate_query = TRUE)
write.csv(PaidUsers.TY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/2_4-PaidUsers-TY.csv",row.names=FALSE,append=FALSE,na="")

# 2_5 - GA - Paid Users - LY
PaidUsers.LY.query.init <- Init(start.date = FYLY.startDate,
                                end.date = FYLY.endDate,
                                dimensions = c(""),
                                metrics = c("ga:users"),
                                filters = "ga:channelGrouping==Paid Search,ga:channelGrouping==Display,ga:source==criteo,ga:medium==paid-social,ga:medium==paid",
                                max.results=10000,
                                table.id = "ga:149139688")
PaidUsers.LY.query <- QueryBuilder(PaidUsers.LY.query.init)
PaidUsers.LY.data <- GetReportData(PaidUsers.LY.query, token, paginate_query = TRUE)
write.csv(PaidUsers.LY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/2_5-PaidUsers-LY.csv",row.names=FALSE,append=FALSE,na="")

#3_1_1 - GA - Organic Users - TW
OrganicUsers.TW.query.init <- Init(start.date = TW.startDate,
                                   end.date = TW.endDate,
                                   dimensions = c("ga:isoYearIsoWeek"),
                                   metrics = c("ga:users"),
                                   filters = "ga:channelGrouping==Organic Search",
                                   max.results=10000,
                                   table.id = "ga:149139688")
OrganicUsers.TW.query <- QueryBuilder(OrganicUsers.TW.query.init)
OrganicUsers.TW.data <- GetReportData(OrganicUsers.TW.query, token, paginate_query = TRUE)

# 3_1_2 - GA - Organic Users - LW
OrganicUsers.LW.query.init <- Init(start.date = LW.startDate,
                                   end.date = LW.endDate,
                                   dimensions = c("ga:isoYearIsoWeek"),
                                   metrics = c("ga:users"),
                                   filters = "ga:channelGrouping==Organic Search",
                                   max.results=10000,
                                   table.id = "ga:149139688")
OrganicUsers.LW.query <- QueryBuilder(OrganicUsers.LW.query.init)
OrganicUsers.LW.data <- GetReportData(OrganicUsers.LW.query, token, paginate_query = TRUE)

# 3_1_3 - GA - Organic Users - TWLY
OrganicUsers.TWLY.query.init <- Init(start.date = "2017-04-01",
                                     end.date = FYTY.endDate,
                                     dimensions = c("ga:isoYearIsoWeek"),
                                     metrics = c("ga:users"),
                                     filters = paste("ga:users>0;ga:channelGrouping==Organic Search;ga:isoYearIsoWeek==",isoWeek.LY,sep=""),
                                     max.results=10000,
                                     table.id = "ga:149139688")
OrganicUsers.TWLY.query <- QueryBuilder(OrganicUsers.TWLY.query.init)
OrganicUsers.TWLY.data <- GetReportData(OrganicUsers.TWLY.query, token, paginate_query = TRUE)

# 3_1 - GA - Organic Users - TW
OrganicUsers.Week.data <- rbind(OrganicUsers.TW.data,OrganicUsers.LW.data,OrganicUsers.TWLY.data)
write.csv(OrganicUsers.Week.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/3_1-OrganicUsers-TW.csv",row.names=FALSE,append=FALSE,na="")

# 3_2 - GA - Organic Users - TM
OrganicUsers.TM.query.init <- Init(start.date = TMTY.startDate,
                                   end.date = TMTY.endDate,
                                   dimensions = c("ga:yearMonth"),
                                   metrics = c("ga:users"),
                                   filters = "ga:channelGrouping==Organic Search",
                                   max.results=10000,
                                   table.id = "ga:149139688")
OrganicUsers.TM.query <- QueryBuilder(OrganicUsers.TM.query.init)
OrganicUsers.TM.data <- GetReportData(OrganicUsers.TM.query, token, paginate_query = TRUE)
write.csv(OrganicUsers.TM.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/3_2-OrganicUsers-TM.csv",row.names=FALSE,append=FALSE,na="")

# 3_3 - GA - Organic Users - TM LY
OrganicUsers.TMLY.query.init <- Init(start.date = TMLY.startDate,
                                     end.date = TMLY.endDate,
                                     dimensions = c("ga:yearMonth"),
                                     metrics = c("ga:users"),
                                     filters = "ga:channelGrouping==Organic Search",
                                     max.results=10000,
                                     table.id = "ga:149139688")
OrganicUsers.TMLY.query <- QueryBuilder(OrganicUsers.TMLY.query.init)
OrganicUsers.TMLY.data <- GetReportData(OrganicUsers.TMLY.query, token, paginate_query = TRUE)
write.csv(OrganicUsers.TMLY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/3_3-OrganicUsers-TMLY.csv",row.names=FALSE,append=FALSE,na="")

# 3_4 - GA - Organic Users - TY
OrganicUsers.TY.query.init <- Init(start.date = FYTY.startDate,
                                   end.date = FYTY.endDate,
                                   dimensions = c(""),
                                   metrics = c("ga:users"),
                                   filters = "ga:channelGrouping==Organic Search",
                                   max.results=10000,
                                   table.id = "ga:149139688")
OrganicUsers.TY.query <- QueryBuilder(OrganicUsers.TY.query.init)
OrganicUsers.TY.data <- GetReportData(OrganicUsers.TY.query, token, paginate_query = TRUE)
write.csv(OrganicUsers.TY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/3_4-OrganicUsers-TY.csv",row.names=FALSE,append=FALSE,na="")

# 3_5 - GA - Organic Users - LY
OrganicUsers.LY.query.init <- Init(start.date = FYLY.startDate,
                                   end.date = FYLY.endDate,
                                   dimensions = c(""),
                                   metrics = c("ga:users"),
                                   filters = "ga:channelGrouping==Organic Search",
                                   max.results=10000,
                                   table.id = "ga:149139688")
OrganicUsers.LY.query <- QueryBuilder(OrganicUsers.LY.query.init)
OrganicUsers.LY.data <- GetReportData(OrganicUsers.LY.query, token, paginate_query = TRUE)
write.csv(OrganicUsers.LY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/3_5-OrganicUsers-LY.csv",row.names=FALSE,append=FALSE,na="")

#4_1_1 - GA - Email Users - TW
EmailUsers.TW.query.init <- Init(start.date = TW.startDate,
                                 end.date = TW.endDate,
                                 dimensions = c("ga:isoYearIsoWeek"),
                                 metrics = c("ga:users"),
                                 filters = "ga:channelGrouping==Email",
                                 max.results=10000,
                                 table.id = "ga:149139688")
EmailUsers.TW.query <- QueryBuilder(EmailUsers.TW.query.init)
EmailUsers.TW.data <- GetReportData(EmailUsers.TW.query, token, paginate_query = TRUE)

# 4_1_2 - GA - Email Users - LW
EmailUsers.LW.query.init <- Init(start.date = LW.startDate,
                                 end.date = LW.endDate,
                                 dimensions = c("ga:isoYearIsoWeek"),
                                 metrics = c("ga:users"),
                                 filters = "ga:channelGrouping==Email",
                                 max.results=10000,
                                 table.id = "ga:149139688")
EmailUsers.LW.query <- QueryBuilder(EmailUsers.LW.query.init)
EmailUsers.LW.data <- GetReportData(EmailUsers.LW.query, token, paginate_query = TRUE)

# 4_1_3 - GA - Email Users - TWLY
EmailUsers.TWLY.query.init <- Init(start.date = "2017-04-01",
                                   end.date = FYTY.endDate,
                                   dimensions = c("ga:isoYearIsoWeek"),
                                   metrics = c("ga:users"),
                                   filters = paste("ga:users>0;ga:channelGrouping==Email;ga:isoYearIsoWeek==",isoWeek.LY,sep=""),
                                   max.results=10000,
                                   table.id = "ga:149139688")
EmailUsers.TWLY.query <- QueryBuilder(EmailUsers.TWLY.query.init)
EmailUsers.TWLY.data <- GetReportData(EmailUsers.TWLY.query, token, paginate_query = TRUE)

# 4_1 - GA - Email Users - TW
EmailUsers.Week.data <- rbind(EmailUsers.TW.data,EmailUsers.LW.data,EmailUsers.TWLY.data)
write.csv(EmailUsers.Week.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/4_1-EmailUsers-TW.csv",row.names=FALSE,append=FALSE,na="")

# 4_2 - GA - Email Users - TM
EmailUsers.TM.query.init <- Init(start.date = TMTY.startDate,
                                 end.date = TMTY.endDate,
                                 dimensions = c("ga:yearMonth"),
                                 metrics = c("ga:users"),
                                 filters = "ga:channelGrouping==Email",
                                 max.results=10000,
                                 table.id = "ga:149139688")
EmailUsers.TM.query <- QueryBuilder(EmailUsers.TM.query.init)
EmailUsers.TM.data <- GetReportData(EmailUsers.TM.query, token, paginate_query = TRUE)
write.csv(EmailUsers.TM.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/4_2-EmailUsers-TM.csv",row.names=FALSE,append=FALSE,na="")

# 4_3 - GA - Email Users - TM LY
EmailUsers.TMLY.query.init <- Init(start.date = TMLY.startDate,
                                   end.date = TMLY.endDate,
                                   dimensions = c("ga:yearMonth"),
                                   metrics = c("ga:users"),
                                   filters = "ga:channelGrouping==Email",
                                   max.results=10000,
                                   table.id = "ga:149139688")
EmailUsers.TMLY.query <- QueryBuilder(EmailUsers.TMLY.query.init)
EmailUsers.TMLY.data <- GetReportData(EmailUsers.TMLY.query, token, paginate_query = TRUE)
write.csv(EmailUsers.TMLY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/4_3-EmailUsers-TMLY.csv",row.names=FALSE,append=FALSE,na="")

# 4_4 - GA - Email Users - TY
EmailUsers.TY.query.init <- Init(start.date = FYTY.startDate,
                                 end.date = FYTY.endDate,
                                 dimensions = c(""),
                                 metrics = c("ga:users"),
                                 filters = "ga:channelGrouping==Email",
                                 max.results=10000,
                                 table.id = "ga:149139688")
EmailUsers.TY.query <- QueryBuilder(EmailUsers.TY.query.init)
EmailUsers.TY.data <- GetReportData(EmailUsers.TY.query, token, paginate_query = TRUE)
write.csv(EmailUsers.TY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/4_4-EmailUsers-TY.csv",row.names=FALSE,append=FALSE,na="")

# 4_5 - GA - Email Users - LY
EmailUsers.LY.query.init <- Init(start.date = FYLY.startDate,
                                 end.date = FYLY.endDate,
                                 dimensions = c(""),
                                 metrics = c("ga:users"),
                                 filters = "ga:channelGrouping==Email",
                                 max.results=10000,
                                 table.id = "ga:149139688")
EmailUsers.LY.query <- QueryBuilder(EmailUsers.LY.query.init)
EmailUsers.LY.data <- GetReportData(EmailUsers.LY.query, token, paginate_query = TRUE)
write.csv(EmailUsers.LY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/4_5-EmailUsers-LY.csv",row.names=FALSE,append=FALSE,na="")

# #5_1_1 - GA - Goal Completions - TW
# GoalCompletions.TW.query.init <- Init(start.date = TW.startDate,
#                                       end.date = TW.endDate,
#                                       dimensions = c("ga:isoYearIsoWeek"),
#                                       metrics = c("ga:goal2Completions","ga:goal3Completions","ga:goal4Completions","ga:goal6Completions","ga:goal8Completions","ga:goal9Completions","ga:goal11Completions"),
#                                       filters ="ga:goal2Completions>0,ga:goal3Completions>0,ga:goal4Completions>0,ga:goal6Completions>0,ga:goal8Completions>0,ga:goal9Completions>0,ga:goal11Completions>0",
#                                       max.results=10000,
#                                       table.id = "ga:149139688")
# GoalCompletions.TW.query <- QueryBuilder(GoalCompletions.TW.query.init)
# GoalCompletions.TW.data <- GetReportData(GoalCompletions.TW.query, token, paginate_query = TRUE)
# 
# # 5_1_2 - GA - Goal Completions - LW
# GoalCompletions.LW.query.init <- Init(start.date = LW.startDate,
#                                       end.date = LW.endDate,
#                                       dimensions = c("ga:isoYearIsoWeek"),
#                                       metrics = c("ga:goal2Completions","ga:goal3Completions","ga:goal4Completions","ga:goal6Completions","ga:goal8Completions","ga:goal9Completions","ga:goal11Completions"),
#                                       filters ="ga:goal2Completions>0,ga:goal3Completions>0,ga:goal4Completions>0,ga:goal6Completions>0,ga:goal8Completions>0,ga:goal9Completions>0,ga:goal11Completions>0",
#                                       max.results=10000,
#                                       table.id = "ga:149139688")
# GoalCompletions.LW.query <- QueryBuilder(GoalCompletions.LW.query.init)
# GoalCompletions.LW.data <- GetReportData(GoalCompletions.LW.query, token, paginate_query = TRUE)
# 
# # 5_1_3 - GA - Goal Completions - TWLY
# GoalCompletions.TWLY.query.init <- Init(start.date = "2017-04-01",
#                                         end.date = FYTY.endDate,
#                                         dimensions = c("ga:isoYearIsoWeek"),
#                                         metrics = c("ga:goal2Completions","ga:goal3Completions","ga:goal4Completions","ga:goal6Completions","ga:goal8Completions","ga:goal9Completions","ga:goal11Completions"),
#                                         filters =paste("ga:goal2Completions>0,ga:goal3Completions>0,ga:goal4Completions>0,ga:goal6Completions>0,ga:goal8Completions>0,ga:goal9Completions>0,ga:goal11Completions>0;",isoWeek.LY,sep=""),
#                                         max.results=10000,
#                                         table.id = "ga:149139688")
# GoalCompletions.TWLY.query <- QueryBuilder(GoalCompletions.TWLY.query.init)
# GoalCompletions.TWLY.data <- GetReportData(GoalCompletions.TWLY.query, token, paginate_query = TRUE)
# 
# # 5_1 - GA - Goal Completions - TW
# GoalCompletions.Week.data <- rbind(GoalCompletions.TW.data,GoalCompletions.LW.data,GoalCompletions.TWLY.data)
# write.csv(GoalCompletions.Week.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/5_1-GoalCompletions-TW.csv",row.names=FALSE,append=FALSE,na="")
# 
# # 5_2 - GA - Goal Completions - TM
# GoalCompletions.TM.query.init <- Init(start.date = TMTY.startDate,
#                                       end.date = TMTY.endDate,
#                                       dimensions = c("ga:yearMonth"),
#                                       metrics = c("ga:goal2Completions","ga:goal3Completions","ga:goal4Completions","ga:goal6Completions","ga:goal8Completions","ga:goal9Completions","ga:goal11Completions"),
#                                       max.results=10000,
#                                       table.id = "ga:149139688")
# GoalCompletions.TM.query <- QueryBuilder(GoalCompletions.TM.query.init)
# GoalCompletions.TM.data <- GetReportData(GoalCompletions.TM.query, token, paginate_query = TRUE)
# write.csv(GoalCompletions.TM.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/5_2-GoalCompletions-TM.csv",row.names=FALSE,append=FALSE,na="")
# 
# # 5_3 - GA - Goal Completions - TM LY
# GoalCompletions.TMLY.query.init <- Init(start.date = TMLY.startDate,
#                                         end.date = TMLY.endDate,
#                                         dimensions = c("ga:yearMonth"),
#                                         metrics = c("ga:goal2Completions","ga:goal3Completions","ga:goal4Completions","ga:goal6Completions","ga:goal8Completions","ga:goal9Completions","ga:goal11Completions"),
#                                         max.results=10000,
#                                         table.id = "ga:149139688")
# GoalCompletions.TMLY.query <- QueryBuilder(GoalCompletions.TMLY.query.init)
# GoalCompletions.TMLY.data <- GetReportData(GoalCompletions.TMLY.query, token, paginate_query = TRUE)
# write.csv(GoalCompletions.TMLY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/5_3-GoalCompletions-TMLY.csv",row.names=FALSE,append=FALSE,na="")
# 
# # 5_4 - GA - Goal Completions - TY
# GoalCompletions.TY.query.init <- Init(start.date = FYTY.startDate,
#                                       end.date = FYTY.endDate,
#                                       dimensions = c(""),
#                                       metrics = c("ga:goal2Completions","ga:goal3Completions","ga:goal4Completions","ga:goal6Completions","ga:goal8Completions","ga:goal9Completions","ga:goal11Completions"),
#                                       max.results=10000,
#                                       table.id = "ga:149139688")
# GoalCompletions.TY.query <- QueryBuilder(GoalCompletions.TY.query.init)
# GoalCompletions.TY.data <- GetReportData(GoalCompletions.TY.query, token, paginate_query = TRUE)
# write.csv(GoalCompletions.TY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/5_4-GoalCompletions-TY.csv",row.names=FALSE,append=FALSE,na="")
# 
# # 5_5 - GA - Goal Completions - LY
# GoalCompletions.LY.query.init <- Init(start.date = FYLY.startDate,
#                                       end.date = FYLY.endDate,
#                                       dimensions = c(""),
#                                       metrics = c("ga:goal2Completions","ga:goal3Completions","ga:goal4Completions","ga:goal6Completions","ga:goal8Completions","ga:goal9Completions","ga:goal11Completions"),
#                                       max.results=10000,
#                                       table.id = "ga:149139688")
# GoalCompletions.LY.query <- QueryBuilder(GoalCompletions.LY.query.init)
# GoalCompletions.LY.data <- GetReportData(GoalCompletions.LY.query, token, paginate_query = TRUE)
# write.csv(GoalCompletions.LY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/5_5-GoalCompletions-LY.csv",row.names=FALSE,append=FALSE,na="")

#6_1_1 - GA - PPC Cost - TW
PPCCost.TW.query.init <- Init(start.date = TW.startDate,
                              end.date = TW.endDate,
                              dimensions = c("ga:isoYearIsoWeek"),
                              metrics = c("ga:adCost"),
                              max.results=10000,
                              table.id = "ga:149139688")
PPCCost.TW.query <- QueryBuilder(PPCCost.TW.query.init)
PPCCost.TW.data <- GetReportData(PPCCost.TW.query, token, paginate_query = TRUE)

# 6_1_2 - GA - PPC Cost - LW
PPCCost.LW.query.init <- Init(start.date = LW.startDate,
                              end.date = LW.endDate,
                              dimensions = c("ga:isoYearIsoWeek"),
                              metrics = c("ga:adCost"),,
                              max.results=10000,
                              table.id = "ga:149139688")
PPCCost.LW.query <- QueryBuilder(PPCCost.LW.query.init)
PPCCost.LW.data <- GetReportData(PPCCost.LW.query, token, paginate_query = TRUE)

# 6_1_3 - GA - PPC Cost - TWLY
PPCCost.TWLY.query.init <- Init(start.date = "2017-04-01",
                                end.date = FYTY.endDate,
                                dimensions = c("ga:isoYearIsoWeek"),
                                metrics = c("ga:adCost"),
                                filters = paste("ga:adCost>0;ga:isoYearIsoWeek==",isoWeek.LY,sep=""),
                                max.results=10000,
                                table.id = "ga:149139688")
PPCCost.TWLY.query <- QueryBuilder(PPCCost.TWLY.query.init)
PPCCost.TWLY.data <- GetReportData(PPCCost.TWLY.query, token, paginate_query = TRUE)

# 6_1 - GA - PPC Cost - TW
PPCCost.Week.data <- rbind(PPCCost.TW.data,PPCCost.LW.data,PPCCost.TWLY.data)
write.csv(PPCCost.Week.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/6_1-PPCCost-TW.csv",row.names=FALSE,append=FALSE,na="")

# 6_2 - GA - PPC Cost - TM
PPCCost.TM.query.init <- Init(start.date = TMTY.startDate,
                              end.date = TMTY.endDate,
                              dimensions = c("ga:yearMonth"),
                              metrics = c("ga:adCost"),
                              max.results=10000,
                              table.id = "ga:149139688")
PPCCost.TM.query <- QueryBuilder(PPCCost.TM.query.init)
PPCCost.TM.data <- GetReportData(PPCCost.TM.query, token, paginate_query = TRUE)
write.csv(PPCCost.TM.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/6_2-PPCCost-TM.csv",row.names=FALSE,append=FALSE,na="")

# 6_3 - GA - PPC Cost - TM LY
PPCCost.TMLY.query.init <- Init(start.date = TMLY.startDate,
                                end.date = TMLY.endDate,
                                dimensions = c("ga:yearMonth"),
                                metrics = c("ga:adCost"),
                                max.results=10000,
                                table.id = "ga:149139688")
PPCCost.TMLY.query <- QueryBuilder(PPCCost.TMLY.query.init)
PPCCost.TMLY.data <- GetReportData(PPCCost.TMLY.query, token, paginate_query = TRUE)
write.csv(PPCCost.TMLY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/6_3-PPCCost-TMLY.csv",row.names=FALSE,append=FALSE,na="")

# 6_4 - GA - PPC Cost - TY
PPCCost.TY.query.init <- Init(start.date = FYTY.startDate,
                              end.date = FYTY.endDate,
                              dimensions = c(""),
                              metrics = c("ga:adCost"),
                              max.results=10000,
                              table.id = "ga:149139688")
PPCCost.TY.query <- QueryBuilder(PPCCost.TY.query.init)
PPCCost.TY.data <- GetReportData(PPCCost.TY.query, token, paginate_query = TRUE)
write.csv(PPCCost.TY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/6_4-PPCCost-TY.csv",row.names=FALSE,append=FALSE,na="")

# 6_5 - GA - PPC Cost - LY
PPCCost.LY.query.init <- Init(start.date = FYLY.startDate,
                              end.date = FYLY.endDate,
                              dimensions = c(""),
                              metrics = c("ga:adCost"),
                              max.results=10000,
                              table.id = "ga:149139688")
PPCCost.LY.query <- QueryBuilder(PPCCost.LY.query.init)
PPCCost.LY.data <- GetReportData(PPCCost.LY.query, token, paginate_query = TRUE)
write.csv(PPCCost.LY.data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/6_5-PPCCost-LY.csv",row.names=FALSE,append=FALSE,na="")

# 7_1 - MCF Goal Completions
mcf_enquiries_data <- get_mcf(profileId = "149139688",
                              start.date = "2017-04-01", end.date = "yesterday",
                              metrics = "mcf:totalConversions",
                              dimensions = "mcf:conversionDate, mcf:basicChannelGroupingPath, mcf:sourceMediumPath",
                              sort = NULL,
                              filters = "mcf:conversionGoalNumber==002,mcf:conversionGoalNumber==003,mcf:conversionGoalNumber==004,mcf:conversionGoalNumber==006,mcf:conversionGoalNumber==008,mcf:conversionGoalNumber==009,mcf:conversionGoalNumber==011",
                              samplingLevel = NULL,
                              start.index = NULL,
                              max.results = NULL,
                              fetch.by = NULL)
write.csv(mcf_enquiries_data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/7_1-MCFGoalCompletions.csv",row.names=FALSE,append=FALSE,na="")

# 7_2 - MCF Phonecalls
mcf_enquiries_data <- get_mcf(profileId = "149139688",
                              start.date = "2017-04-01", end.date = "yesterday",
                              metrics = "mcf:totalConversions",
                              dimensions = "mcf:conversionDate, mcf:basicChannelGroupingPath, mcf:sourceMediumPath",
                              sort = NULL,
                              filters = "mcf:conversionGoalNumber==003",
                              samplingLevel = NULL,
                              start.index = NULL,
                              max.results = NULL,
                              fetch.by = NULL)
write.csv(mcf_enquiries_data,"V:/Shared Documents/ScoreCard/GoogleAnalytics/Output/7_2-MCFPhonecalls.csv",row.names=FALSE,append=FALSE,na="")
