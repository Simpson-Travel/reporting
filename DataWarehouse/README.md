#DW_RunCheck

##Process to check if the Data Warehouse has updated overnight.
DW_RunCheck.xlsm has two queries built in Power Query that look up against log.VersioningStats.
23 tables should be built each day.
If all have been built, then Run.Today=TRUE and the TablesCheck return will be empty
If less than 23 tables have been built, then Run.Today=FALSE and the TablesCheck will return all tables that have not been succesfully built today.
DW_RunCheck.vbs is run via Task Scheduler at 6am each morning to refresh the file.

##DataWarehouse Status
DataWarehouseStatus.r reads data from DW_RunCheck.xlsm.
If file has been refreshed today, and all tables have been built, then no action is taken.
If file has not been refreshed today, an email is sent as something has broken.
If file has been refreshed today and not all tables have been built, an email is sent listing the failed tables.
Runs via Task Scheduler at 6:05am each morning.
