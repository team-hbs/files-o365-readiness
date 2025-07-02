param (
    [int]$batchNumber = -1,
	[string]$wave = '',
    [int]$sourceId = -1
)

. .\common.ps1


$sources = GetSources -batchnumber $batchNumber -wave $wave -ownerId $sourceId
foreach($source in $sources)
{
    $path = $source.ADHomeDirectory
    write-host 'Getting File count for' $path '...'
    #$overallFileCount = (Get-ChildItem -LiteralPath $path -Recurse -File | Measure-Object | Select-Object -Property Count).Count
    #$ownerId = $source.Id
    #$query = "UPDATE Source SET FileCount = " + $overallFileCount + " WHERE Id = " + $ownerId
    #SqlQueryInsert -query $query
	
	
	$overall = Get-ChildItem $path -recurse | Measure-Object -property length -sum
	$overallFileSize = $overall.Sum / 1000000000
	$overallFileCount = $overall.Count
	
	$query = "UPDATE Source SET FileCount = " + $overallFileCount + ", FileSize = " + $overallFileSize + " WHERE Id = " + $ownerId
    SqlQueryInsert -query $query
	
}