write-host 'Do not run this file!' -f Yellow
Exit

#unblock files
set-executionpolicy unrestricted
unblock-file -path .\pre_migration_master.ps1
unblock-file -path .\crawl_v11.ps1
unblock-file -path .\common.ps1
unblock-file -path .\installAzCopy.ps1
unblock-file -path .\reports.ps1
unblock-file -path .\OfficeMonitor.ps1
#install modules
.\pre_migration_master.ps1 -mode 'Install' 

<# For SQL Server Mode
.\pre_migration_master.ps1 -mode 'SetConfig' -key 'DatabaseMode' -value 'SQLServer'
.\pre_migration_master.ps1 -mode 'SetConfig' -key 'DatabaseServer' -value 'HBS-JBALDWIN01'
.\pre_migration_master.ps1 -mode 'SetConfig' -key 'DatabaseName' -value 'Migration'
.\pre_migration_master.ps1 -mode 'CreateDatabase' 
#>

.\pre_migration_master.ps1 -mode 'Import' 
.\pre_migration_master.ps1 -mode 'Scan' -SourceId 11 #run scan on source id (owner id)
.\pre_migration_master.ps1 -mode 'Scan' -BatchNumber 0 #run scan on batch number
.\pre_migration_master.ps1 -mode 'Scan' #run scan on all sources
.\reports.ps1 #generate reports on all sources
.\reports.ps1 -SourceId 11  #generate reports on source id (owner id)
.\reports.ps1 -BatchNumber 0 #generate reports on batch number


.\pre_migration_master.ps1 -mode 'SetConfig' -key 'NoOffice' -value 'true'