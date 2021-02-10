param (
	[int] $ownerId = -1,
	[int] $batchNumber = -1
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

$global:DataSource = $PSScriptRoot + "\FileToOneDrive.db"
$global:SqlServer = $false



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


$sqlOverall = "SELECT Id, 
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
	FROM Files_Batch_users,Files_Users 
	WHERE Files_Batch_users.Id = Files_Users.OwnerId
	AND	Files_Users.OwnerId = {OWNER_ID}"
	
$sqlTopLevelFolders = "SELECT DISTINCT Folder01, Folder02, Folder03
FROM Files_OneDrive
WHERE OwnerId = {OWNER_ID}"

$sqlOverGB = "SELECT ParentFolder, FileName, Size 
FROM Files_OneDrive
WHERE Size > 2000
AND OwnerID = {OWNER_ID}
ORDER BY size desc"

$sqlErrors = "SELECT FileName,Extension,PathLength,ParentFolder,HasMacro,Error,Size,Created,Modified
FROM Files_OneDrive
WHERE Error IS NOT NULL
AND Error != ' '
AND Error != ''
AND OwnerId = {OWNER_ID}
"

$sqlPathLengthExceeded =  "SELECT FileName,Extension,PathLength,ParentFolder,HasMacro,Error,Size,Created,Modified
FROM Files_OneDrive
WHERE PathLength > 218
AND OwnerId = {OWNER_ID}
"

$sqlExtensionCount = "SELECT Extension, Count(Extension) as Total
FROM Files_OneDrive
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
	$batchQuery = "SELECT * FROM Files_Batch_users WHERE BatchNumber = $batchNumber"
	$result = SqlQueryReturn($batchQuery)
}
elseif($ownerId -ne -1)
{
	$ownerQuery = "SELECT * FROM Files_Batch_users WHERE Id = $ownerId"
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
                FROM Files_OneDrive
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
