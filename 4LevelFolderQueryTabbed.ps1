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

$LogCounts = [System.Collections.Generic.List[Object]]::new()

#$global:DataSource = $PSScriptRoot + "\FileToOneDrive.db"

function GetConfig($key) {
	$value = $null
	$null = @(
		$query = "SELECT * FROM Config WHERE Key='" + $key + "'"
		#write-host $query -f yellow
		$rows = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
		$value = $rows[0].Value 
	)
	if ($rows[0].Encrypted -eq 1) {
		$Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($value)
		$result = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
		[System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr)
		$value = $result
	}
	return $value
}

function SetConfig($key, $value, $encrypted) {
	$null = @(
		if ($encrypted) {
			$value = ConvertTo-SecureString -String [string] $value -AsPlainText -Force 
			$encrypted = 1
		}
		else {
			$encrypted = 0
		}
		$query = "UPDATE Config SET Value='" + $value + "',Encrypted=" + $encrypted + " WHERE Key = '" + $key + "'"
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	)
}

$global:DataSource = $PSScriptRoot + "\FilesToO365.db"
$global:SqlServer = $false

if ((GetConfig 'DatabaseMode') -eq 'SQLServer') {
	$connectionString = GetConfig 'ConnectionString'
	$databaseServer = GetConfig 'DatabaseServer'
	$databaseName = GetConfig 'DatabaseName'
	if ($connectionString.Trim() -eq '') {
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

#Queries for other tabs
#Overall Query
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
	
#Top Level Folder
$sqlTopLevelFolders = "SELECT DISTINCT Folder01, Folder02, Folder03, Folder04
FROM ScanFile
WHERE OwnerId = {OWNER_ID}"

#Over2gb Query
$sqlOverGB = "SELECT ParentFolder, FileName, Size 
FROM ScanFile
WHERE Size > 2000
AND OwnerID = {OWNER_ID}
ORDER BY size desc"

#Error Query
$sqlErrors = "SELECT FileName,Extension,PathLength,ParentFolder,HasMacro,Error,Size,Created,Modified
FROM ScanFile
WHERE Error IS NOT NULL
AND Error != ' '
AND Error != ''
AND OwnerId = {OWNER_ID}
"

#Path length Query
$sqlPathLengthExceeded = "SELECT FileName,Extension,PathLength,ParentFolder,HasMacro,Error,Size,Created,Modified
FROM ScanFile
WHERE PathLength > 218
AND OwnerId = {OWNER_ID}
"

#Extension Query
$sqlExtensionCount = "SELECT Extension, Count(Extension) as Total
FROM ScanFile
WHERE OwnerId = {OWNER_ID}
GROUP By Extension
ORDER BY Extension ASC
"

#Runs reports for each worksheet tab
$reports = @(
	@{Query = $sqlOverall; Name = 'Overall' }
	, @{Query = $sqlOverGB; Name = 'Over 2GB' }
	, @{Query = $sqlErrors; Name = 'Errors' }
	, @{Query = $sqlPathLengthExceeded; Name = 'Path Length Too Long' }
	, @{Query = $sqlExtensionCount; Name = 'Extensions' }
)

if ($batchNumber -ne -1) {
	$batchQuery = "SELECT * FROM Source WHERE BatchNumber = $batchNumber"
	$result = SqlQueryReturn($batchQuery)
}
elseif ($ownerId -ne -1) {
	$ownerQuery = "SELECT * FROM Source WHERE Id = $ownerId"
	$result = SqlQueryReturn($ownerQuery)
}
else {
	$ownerQuery = "SELECT * FROM Source ORDER BY id"
	$result = SqlQueryReturn($ownerQuery)
}


foreach ($row in $result) {	

	#Creates $directory Path		
	$directory = $row.ADHomeDirectory

	#Replaced directory patch name for $logfile normalized name
	$Pathdirectory = $directory.Replace('\\', '')
	$Pathdirectory = $Pathdirectory.Replace('\', '_')
	$Pathdirectory = $Pathdirectory.Replace(':', '')
	$Pathdirectory = $Pathdirectory.Replace('.', '_')
	$logFile = '.\' + $Pathdirectory + '_scan.xlsx'
	$ownerId = $row.Id

	#Creates new worksheet tab for each report query
	foreach ($report in $reports) {
		$query = $report.Query
		$query = $query.Replace('{OWNER_ID}', $ownerId)
		$result = SqlQueryReturn($query)
		$reportName = $report.Name
		if ($report.Name -eq 'Overall') {
			if ($result.Extensions.Length -gt 2999) {
				$result.Extensions = $result.Extensions.SubString(0, 3000)
			}
			$unixEpoch = Get-Date -Date "01/01/1970"
			#$result.CreatedDate =  $unixEpoch.AddSeconds($result.CreatedDate)
		}
		#exports report queries to excel file
		$result | Select * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | Export-Excel $logFile -WorksheetName $reportName -AutoSize -MaxAutoSizeRows 2 
	}
	
	#Gets top level folders for each locations (2 levels down - Add more folders for more levels down)
	$query = $sqlTopLevelFolders.Replace('{OWNER_ID}', $ownerId)
	$result = SqlQueryReturn($query)
	$topLevelFolders = @()
	foreach ($row in $result) {
		if ($row.Folder01 -ne $null -AND $row.Folder02 -ne $null -AND $row.Folder03 -ne $null -AND $row.Folder04 -ne $null) {
			$path = '\\' + $row.Folder01 + '\' + $row.Folder02 + '\' + $row.Folder03 + '\' + $row.Folder04

			#Query for file count of each folder under the parent folder
			$sqlCounts = "SELECT Count(Id) as FileCount, Sum(Size) as FileSize
                FROM ScanFile
                WHERE Path like '{PARENT_FOLDER}%'
                AND Folder01 = '" + ($row.Folder01).Replace("'", "''") + "'
                AND Folder02 = '" + ($row.Folder02).Replace("'", "''") + "'
                AND Folder03 = '" + ($row.Folder03).Replace("'", "''") + "'
				AND Folder04 = '" + ($row.Folder04).Replace("'", "''") + "'"
            
			#Runs top level folder query
			$query = $sqlCounts.Replace('{PARENT_FOLDER}',$directory)
			$totals = SqlQueryReturn($query)

			#For each query returned - appends to $logCounts
			foreach ($total in $totals) {


                $tempItem = New-Object -TypeName PsObject
                $tempItem | Add-Member -MemberType NoteProperty -Name 'Server Name' -Value $row.Folder01
                $tempItem | Add-Member -MemberType NoteProperty -Name 'Share Location' -Value $row.Folder02
                $tempItem | Add-Member -MemberType NoteProperty -Name 'Top Level Folder' -Value $row.Folder03
                $tempItem | Add-Member -MemberType NoteProperty -Name 'Second Level Folder' -Value $row.Folder04
                $tempItem | Add-Member -MemberType NoteProperty -Name 'Path' -Value $path
                $tempItem | Add-Member -MemberType NoteProperty -Name 'File Count' -Value $total.FileCount    
                $tempItem | Add-Member -MemberType NoteProperty -Name 'Size' -Value $total.FileSize
                $logCounts.Add($tempItem)

			}
		}
	}
#Exports $logCounts and appends to reports
$logCounts | Select * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | Export-Excel $logFile -WorksheetName 'Top Level Folders' -AutoSize -MaxAutoSizeRows 2 

}