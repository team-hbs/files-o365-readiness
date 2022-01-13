param (
	[int] $OwnerId = -1,
	[int] $BatchNumber = -1,
    [string] $batches = ''
)


if ((Get-Module -ListAvailable -Name PSSQLite) -ne $null) {
    Import-Module -Name PSSQLite
} 
else {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-Module -Name PSSQLite
}

# Module for interacting with xlsx files
if ((Get-Module -ListAvailable -Name ImportExcel) -ne $null) {
    Import-Module -Name ImportExcel
} 
else {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-Module -Name ImportExcel
}


#$global:DataSource = $PSScriptRoot + "\FileToOneDrive.db"

function GetConfig($key)
{
	$value = $null
	$null = @(
        $query = "SELECT * FROM Config WHERE Key='" + $key + "'"
		#write-host $query -f yellow
        $rows =  Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
		$value = $rows[0].Value 
	)
	if ($rows[0].Encrypted -eq 1)
	{
		$Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($value)
		$result = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
		[System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr)
		$value = $result
	}
	return $value
}

function SetConfig($key, $value, $encrypted)
{
	$null = @(
		if ($encrypted)
		{
			$value = ConvertTo-SecureString -String [string] $value -AsPlainText -Force 
			$encrypted = 1
		}
		else
		{
			$encrypted = 0
		}
	    $query = "UPDATE Config SET Value='" + $value + "',Encrypted=" + $encrypted + " WHERE Key = '" + $key + "'"
        Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	)
}

$global:DataSource = $PSScriptRoot + "\FilesToO365.db"
$global:SqlServer = $false

if ((GetConfig 'DatabaseMode') -eq 'SQLServer')
{
	$connectionString = GetConfig 'ConnectionString'
	$databaseServer = GetConfig 'DatabaseServer'
	$databaseName = GetConfig 'DatabaseName'
	if ($connectionString.Trim() -eq '')
	{
		 $connectionString = 'Server=' + $databaseServer + ';Database=' + $databaseName + ';Integrated Security=true;'
	}
	$global:SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$global:SqlConnection.ConnectionString = $connectionString
	$global:SqlServer = $true
}


function SqlQueryReturn($query) {
	write-host $query -f Yellow
	if ($global:SqlServer) {
		$DataTable = $null
		$null = @(
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd.CommandText = $query
			$SqlCmd.Connection = $global:SqlConnection
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCmd
			$DataSet = New-Object System.Data.DataSet
			$SqlAdapter.Fill($DataSet)
			$DataTable = $DataSet.Tables[0]
		)
		return $DataTable
	} else {
		return Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	}
}


$sqlOverall = "SELECT ScanJob.Id, 
	SamAccountName,
	ADHomeDirectory,
	FileCountDisk,
	FileCountCrawl,
	MacroCount,
    FileSizeDisk,
	FileSizeCrawl,
	ErrorCount,
	OfficeErrorCount,
	OldOfficeCount,
	PathLengthCount, 
	NoAccessCount, 
	CreatedDate
	FROM Source,ScanJob 
	WHERE Source.Id = ScanJob.OwnerId
	AND	ScanJob.OwnerId = {OWNER_ID}"
	
$sqlTopLevelFolders = "SELECT DISTINCT Folder01, Folder02, Folder03
FROM ScanFile
WHERE OwnerId = {OWNER_ID}"

$sqlOverGB = "SELECT ParentFolder, FileName, Size 
FROM ScanFile
WHERE Size > 2000
AND OwnerID = {OWNER_ID}
ORDER BY size desc"

$sqlErrors = "SELECT FileName,Extension,PathLength,ParentFolder,HasMacro,Error,Size,Created,Modified
FROM ScanFile
WHERE Error IS NOT NULL
AND Error != ' '
AND Error != ''
AND OwnerId = {OWNER_ID}
"

$sqlPathLengthExceeded =  "SELECT FileName,Extension,PathLength,ParentFolder,HasMacro,Error,Size,Created,Modified
FROM ScanFile
WHERE PathLength > 218
AND OwnerId = {OWNER_ID}
"

$sqlExtensionCount = "SELECT Extension, Count(Extension) as Total
FROM ScanFile
WHERE OwnerId = {OWNER_ID}
GROUP By Extension
ORDER BY Extension ASC
"

$reports = @(
			@{Query= $sqlOverall; Name ='Overall'}
			,@{Query = $sqlOverGB; Name = 'Over 2GB'}
			,@{Query = $sqlErrors; Name = 'Errors'}
            ,@{Query = $sqlPathLengthExceeded; Name = 'Path Length Too Long'}
            ,@{Query = $sqlExtensionCount; Name = 'Extensions'}
			)

if ($batchNumber -ne -1)
{
	$batchQuery = "SELECT * FROM Source WHERE BatchNumber = $batchNumber"
    $logFile = '.\batch_' + $batchNumber + '_scan.xlsx'
	$sources = SqlQueryReturn($batchQuery)
}
if ($batches -ne '' -AND $batches -ne $null)
{
    $batchQuery = "SELECT * FROM Source WHERE BatchNumber IN (" + $batches + ")"
    $logFile = '.\batch_' + $batches.Replace(',','_') + '_scan.xlsx'
	$sources = SqlQueryReturn($batchQuery)
}
else
{
	$ownerQuery = "SELECT * FROM Source ORDER BY id"
	$sources = SqlQueryReturn($ownerQuery)
    $logFile = '.\all_scan.xlsx'
}
foreach($report in $reports)
{
    $allResults = @()
    foreach($source in $sources)
    {	
	   
	    $ownerId = $source.Id
		$query = $report.Query
		$query = $query.Replace('{OWNER_ID}',$ownerId)
		$result = SqlQueryReturn($query)
		$reportName = $report.Name
		if ($report.Name -eq 'Overall')
		{
			if ($result.Extensions.Length -gt 2999)
			{
				$result.Extensions = $result.Extensions.SubString(0,3000)
			}
			$unixEpoch = Get-Date -Date "01/01/1970"
			$result.CreatedDate =  $result.CreatedDate
		}
        foreach($row in $result)
        {
            $allResults = $allResults + $row
        }
    }
    $allResults |Select * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | Export-Excel $logFile -WorksheetName $reportName -AutoSize -MaxAutoSizeRows 2 
}
   