library(openxlsx)
DW_RunCheck <- read.xlsx('V:/Shared Documents/DataWarehouse/DW_RunCheck.xlsm',sheet=1)
if ((DW_RunCheck[1]==1) & (floor(DW_RunCheck[2])==as.numeric(as.Date(Sys.Date())-as.Date("1900-01-01"))+2)) {
  # IF DATA WAREHOUSE HAS REFRESHED TODAY, FUNCTION HERE WILL RUN
  x=1
}