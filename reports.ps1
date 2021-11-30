param (
	[int] $OwnerId = -1,
	[int] $BatchNumber = -1
)


$commonPath = $PSScriptRoot + "\common.ps1"
. $commonPath

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
	$result = SqlQueryReturn($batchQuery)
}
elseif($ownerId -ne -1)
{
	$ownerQuery = "SELECT * FROM Source WHERE Id = $ownerId"
	$result = SqlQueryReturn($ownerQuery)
}
else
{
	$ownerQuery = "SELECT * FROM Source ORDER BY id"
	$result = SqlQueryReturn($ownerQuery)
}

foreach($row in $result)
{	
    $directory = $row.ADHomeDirectory
	$directory = $directory.Replace('\\','')
	$directory = $directory.Replace('\','_')
	$directory = $directory.Replace(':','')
	$directory = $directory.Replace('.','_')
	$logFile = '.\' + $directory + '_scan.xlsx'
	$ownerId = $row.Id
	foreach($report in $reports)
	{
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
			$result.CreatedDate =  $unixEpoch.AddSeconds($result.CreatedDate)
		}
		$result |Select * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | Export-Excel $logFile -WorksheetName $reportName -AutoSize -MaxAutoSizeRows 2 
	}
    $query = $sqlTopLevelFolders.Replace('{OWNER_ID}',$ownerId)
	$result = SqlQueryReturn($query)
    $topLevelFolders = @()
	foreach($row in $result)
	{
        if ($row.Folder01 -ne $null -AND $row.Folder02 -ne $null -AND $row.Folder03 -ne $null)
        {
            $path = '\\' + $row.Folder01 + '\' + $row.Folder02 + '\' + $row.Folder03
            $query = "SELECT Count(Id) as FileCount, Sum(Size) as FileSize
                FROM ScanFile
                WHERE OwnerId = $ownerId
                AND Folder01 = '" + $row.Folder01 + "'
                AND Folder02 = '" + $row.Folder02 + "'
                AND Folder03 = '" + $row.Folder03 + "'"
            
            $totals = SqlQueryReturn($query)
            foreach($total in $totals)
            {
                $tempItem = New-Object -TypeName PsObject
		        $tempItem | Add-Member -MemberType NoteProperty -Name 'Path' -Value $path
                $tempItem | Add-Member -MemberType NoteProperty -Name 'FileCount' -Value $total.FileCount
                $tempItem | Add-Member -MemberType NoteProperty -Name 'Size' -Value $total.FileSize
                $topLevelFolders = $topLevelFolders + $tempitem
            }
        }
	}
    $topLevelFolders | Select * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | Export-Excel $logFile -WorksheetName 'Top Level Folders' -AutoSize -MaxAutoSizeRows 2 
}
