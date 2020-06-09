<#
param (
		        [string] $startDirectory = ""
      )
#>

# Display window to browse for folder
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$FileBrowser.ShowDialog()
$startDirectory = $FileBrowser.SelectedPath


Install-Module PSSQLite
Import-Module PSSQLite
#Get-Command -Module PSSQLite

$global:DataSource = $PSScriptRoot + "\FileToOneDrive.db"

$crawlPath = $PSScriptRoot + "\crawl_v10.ps1"
. $crawlPath
#. .\office_cleanup.ps1

function GetUsers($batchNumber)
{
	$users = $null
	write-host $batchNumber
	$query = 'SELECT Files_Batch_Users.ADhomeDirectory,Files_Batch_Users.Id FROM Files_Batch_Users WHERE Files_Batch_Users.BatchNumber = ' + $batchNumber + ''
	write-host "Query:" $query -ForegroundColor Green
	$users = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	write-host "running query"
	#write-host $users
	return $users
}

function GetOldOfficeDocuments($directory)
{
	$documents = $null
	$query = "SELECT Files_OneDrive.Path as Path,Files_OneDrive.Id as Id,Files_OneDrive.OwnerId as OwnerId FROM Files_OneDrive, Files_Batch_Users WHERE (Extension = 'xls' OR Extension = 'doc' OR Extension = 'ppt') AND Files_Batch_Users.Id = Files_OneDrive.OwnerID AND Files_Batch_Users.ADhomeDirectory = '" + $directory + "'"
	write-host "Query:" $query -ForegroundColor Green
	$documents = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	write-host "running query"
	#write-host $documents
	return $documents
}

function GetNewBatch($directory)
{
	$users = $null
	write-host $batchNumber
	$query = "SELECT Files_Batch_Users.ADhomeDirectory,Files_Batch_Users.Id FROM Files_Batch_Users WHERE Files_Batch_Users.ADhomeDirectory = '" + $directory + "'"
	write-host "Query:" $query -ForegroundColor Green
	$users = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	write-host "running query"
	#write-host $users
	return $users
}

function CreateNewDirectoryEntry($directory)
{
 
	$query = "Insert INTO  Files_Batch_Users  (SamAccountName, ADhomeDirectory,BatchNumber) Values('jbaldwin', '$directory',0)"
    #write-host "Query:" $query -ForegroundColor Green 	   
	Invoke-SqliteQuery -Query $Query -DataSource $global:DataSource
}

function InitPreMigrationMaster($directory)
{
    #$users = GetUsers $batchNumber
	$users = GetNewBatch $directory
	foreach($row in $users)
	{
		write-host $row -ForegroundColor Cyan
		$path = $row.ADhomedirectory
		$ownerId = $row.Id
		
		#write-host "OwnerId:" $ownerId -ForegroundColor Yellow
		#$email = "jbaldwin@hbs.net"
		ClearCrawlData $ownerId
		$timestamp =  get-date -f _MM_dd_HH_mm_ss
		$logFile = $PSScriptRoot + "\crawl_" + $timestamp + "$ownerId.csv"
		#InitCrawl  $ownerId $email $path $false |  select-object FileName,Message,ParentFolderCurrent, Query  | Export-Csv $logFile
		InitCrawl  $ownerId $email $path $false |  select-object FileName,Message,ParentFolderCurrent, Query  | Export-Csv $logFile
	}
}

function ClearCrawlData($ownerId)
{
    try
    {
		$query = 'DELETE FROM Files_Users Where OwnerId = ' + $ownerId + ''
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
		$query = 'DELETE FROM Files_OneDrive Where OwnerId = ' + $ownerId + ''
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
    }
	catch
	{
    	$line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
		write-error $message
	}
}

function OfficeConversionTest()
{
	#get all old office documents
	$documents = GetOldOfficeDocuments $directory
	foreach($row in $documents)
	{
		write-host $row -ForegroundColor Cyan
		$path = $row.Path
		$ownerId = $row.OwnerId
		$id = $row.Id
		
		#write-host "OwnerId:" $ownerId -ForegroundColor Yellow

		$timestamp =  get-date -f _MM_dd_HH_mm_ss
		$logFile = $PSScriptRoot + "\crawl_" + $timestamp + "$ownerId.csv"
		$result = $null
		$convertMessage = ""
		$file = Get-ChildItem -LiteralPath $path -File -ErrorAction Stop
		try
		{
         $result = ConvertDocument $path $file $saveAs
		}
		catch
		{
			$result = New-Object -TypeName psobject 
			$result | Add-Member -MemberType NoteProperty -Name HasMacro -Value $false
			$result | Add-Member -MemberType NoteProperty -Name ConvertMessage -Value $line.ToString() + ":" + $_.Exception.Message
			$result | Add-Member -MemberType NoteProperty -Name ConvertSuccess -Value $false
		}	
		$convertSuccess = $result.ConvertSuccess
		$convertMessage = $result.ConvertMessage
		$hasMacroValue = 0
		$convertSuccessValue = 0

		if ($result.HasMacro -eq $true)
		{
			$hasMacroValue = 1
		}
		if ($result.ConvertSuccess -eq $true)
		{
			$convertSuccessValue = 1
		}
		#TODO - UPDATE DATABASE ROW
	}

}

function GeneratePostScanReport ($directory)
{
	write-host $global:DataSource
	$users = $null
	write-host $batchNumber
	$query = "SELECT Id,SamAccountName,BatchNumber,ADHomeDirectory,OwnerId,FileCountDisk,FileCountCrawl,MacroCount,Extensions,FileSizeDisk,FileSizeCrawl,ErrorCount,OfficeErrorCount,OldOfficeCount,PathLengthCount,NoAccessCount,CreatedDate FROM Files_Batch_Users,Files_Users WHERE Files_Batch_Users.Id = Files_Users.OwnerId AND ADHomeDirectory = '" + $directory + "'"
	write-host "Query:" $query -ForegroundColor Green
	$report = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	write-host "running query"
	#write-host $users
	$timestamp =  get-date -f _MM_dd_HH_mm_ss
	$logFile = $PSScriptRoot + "\report_" + $timestamp + "$ownerId.csv"
	$report  | Export-Csv $logFile
	

	#Microsoft errors query
	$query = "SELECT OwnerId,SamAccountName,BatchNumber,ADHomeDirectory,FileName,Extension,Path,Error
			` FROM Files_Batch_Users, Files_OneDrive
			` WHERE ADHomeDirectory = '$directory' AND Files_Batch_Users.id = Files_OneDrive.OwnerId AND Error <> '' AND Error <> ' '"
	write-host "Query:" $query -ForegroundColor Green
	$report = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	write-host "running query"
	$timestamp =  get-date -f _MM_dd_HH_mm_ss
	$logFile = $PSScriptRoot + "\msft_errors_" + $timestamp + "$ownerId.csv"
	$report  | Export-Csv $logFile
}

CreateNewDirectoryEntry $startDirectory
InitPreMigrationMaster $startDirectory
# Read-Host -Prompt "Enter to Generate Post Scan Report"
GeneratePostScanReport $startDirectory

# OLD!! usage run as admin --> .\pre_migration_master.ps1 -startDirectory "c:\test"