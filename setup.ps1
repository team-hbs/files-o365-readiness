set-executionpolicy unrestricted
unblock-file -path .\pre_migration_master.ps1
unblock-file -path .\crawl_v11.ps1
unblock-file -path .\create_db.ps1


<# Uncomment for SQL Server Mode
.\pre_migration_master.ps1 -mode 'SetConfig' -key 'DatabaseMode' -value 'SQLServer'
.\pre_migration_master.ps1 -mode 'SetConfig' -key 'DatabaseServer' -value 'HBS-JBALDWIN01'
.\pre_migration_master.ps1 -mode 'SetConfig' -key 'DatabaseName' -value 'Migration'
.\pre_migration_master.ps1 -mode 'CreateDatabase' 
#>


.\pre_migration_master.ps1 -mode 'Scan' -path '.\ad_batch_template.csv' #import AND scan
.\pre_migration_master.ps1 -mode 'Import' -path '.\ad_batch_template.csv' #import only
.\pre_migration_master.ps1 -mode 'Scan' -OwnerId 11 #run scan on source id (owner id)
.\pre_migration_master.ps1 -mode 'Scan' -BatchNumber 0 #run scan on batch number
.\pre_migration_master.ps1 -mode 'Scan' #run scan on all sources
.\reports.ps1 #generate reports on all sources
.\reports.ps1 -OwnerId 11 #generate reports on batch numner
.\reports.ps1 -BatchNumber 0 #generate reports on source id (owner id)

