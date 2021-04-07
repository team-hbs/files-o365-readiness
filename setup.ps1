set-executionpolicy unrestricted
unblock-file -path .\pre_migration_master.ps1
.\pre_migration_master.ps1 -mode 'Install' 

<# Uncomment for SQL Server Mode
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

