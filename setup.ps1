.\migration_master.ps1 -mode 'SetConfig' -key 'Notifications' -value 'false'


<# Uncomment for SQL Server Mode
.\pre_migration_master.ps1 -mode 'SetConfig' -key 'DatabaseMode' -value 'SQLServer'
.\pre_migration_master.ps1 -mode 'SetConfig' -key 'DatabaseServer' -value 'Dc0-sharegate-01'
.\pre_migration_master.ps1 -mode 'SetConfig' -key 'DatabaseName' -value 'PDrive2'
.\pre_migration_master.ps1 -mode 'CreateDatabase' 
#>

#.\pre_migration_master.ps1 -mode 'Scan' -path '.\import\Batch_02.01.21.csv' #import and scan
.\pre_migration_master.ps1 -mode 'Import' -path '.\import\Batches_02_02_21.csv' #import only
.\pre_migration_master.ps1 -mode 'Scan' -ownerId 11 #run scan


