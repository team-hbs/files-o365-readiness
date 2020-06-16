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

# Module for interacting with xlsx files
Install-Module ImportExcel
Import-Module ImportExcel

$global:DataSource = $PSScriptRoot + "\FileToOneDrive.db"

# Create the path to the crawl script and run it
$crawlPath = $PSScriptRoot + "\crawl_v10.ps1"
. $crawlPath
#. .\office_cleanup.ps1

function GetUsers($batchNumber)
{
	$users = $null
	Write-Host $batchNumber
	$query = "	SELECT Files_Batch_Users.ADhomeDirectory, Files_Batch_Users.Id 
				FROM Files_Batch_Users 
				WHERE Files_Batch_Users.BatchNumber = '$batchNumber'"
	Write-Host "Query:" $query -ForegroundColor Green
	$users = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	Write-Host "running query"
	#Write-Host $users
	return $users
}

function GetOldOfficeDocuments($directory)
{
	$documents = $null
	$query = "	SELECT Files_OneDrive.Path as Path, Files_OneDrive.Id as Id, Files_OneDrive.OwnerId as OwnerId 
				FROM Files_OneDrive, Files_Batch_Users 
				WHERE (Extension = 'xls' OR Extension = 'doc' OR Extension = 'ppt') 
				AND Files_Batch_Users.Id = Files_OneDrive.OwnerID AND Files_Batch_Users.ADhomeDirectory = '$directory'"
	Write-Host "Query:" $query -ForegroundColor Green
	$documents = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	Write-Host "running query"
	#Write-Host $documents
	return $documents
}

function GetNewBatch($directory)
{
	$users = $null
	Write-Host $batchNumber
	$query = "	SELECT Files_Batch_Users.ADhomeDirectory, Files_Batch_Users.Id 
				FROM Files_Batch_Users 
				WHERE Files_Batch_Users.ADhomeDirectory = '$directory'"
	Write-Host "Query:" $query -ForegroundColor Green
	$users = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	Write-Host "running query"
	#Write-Host $users
	return $users
}

function CreateNewDirectoryEntry($directory)
{
 
	$query = "Insert INTO Files_Batch_Users (SamAccountName, ADhomeDirectory,BatchNumber) VALUES ('jbaldwin', '$directory', 0)"
    #Write-Host "Query:" $query -ForegroundColor Green 	   
	Invoke-SqliteQuery -Query $Query -DataSource $global:DataSource
}

function InitPreMigrationMaster($directory)
{
    #$users = GetUsers $batchNumber
	$users = GetNewBatch $directory
	foreach($row in $users)
	{
		Write-Host $row -ForegroundColor Cyan
		$path = $row.ADhomedirectory
		$ownerId = $row.Id
		
		#Write-Host "OwnerId:" $ownerId -ForegroundColor Yellow
		ClearCrawlData $ownerId
		$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
		$logFile = $PSScriptRoot + "\crawl_" + $timestamp + "$ownerId.csv"
		InitCrawl  $ownerId $path $false |  Select-Object FileName,Message,ParentFolderCurrent, Query  | Export-Csv $logFile
	}
}

function ClearCrawlData($ownerId)
{
    try
    {
		$query = "DELETE FROM Files_Users WHERE OwnerId = '$ownerId'"
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
		$query = "DELETE FROM Files_OneDrive WHERE OwnerId = '$ownerId'"
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
    }
	catch
	{
    	$line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
		Write-Error $message
	}
}

function OfficeConversionTest()
{
	#get all old office documents
	$documents = GetOldOfficeDocuments $directory
	foreach($row in $documents)
	{
		Write-Host $row -ForegroundColor Cyan
		$path = $row.Path
		$ownerId = $row.OwnerId
		$id = $row.Id
		
		#Write-Host "OwnerId:" $ownerId -ForegroundColor Yellow

		$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
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
	Write-Host $global:DataSource
	$users = $null
	$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
					FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate 
				FROM Files_Batch_Users,Files_Users 
				WHERE Files_Batch_Users.Id = Files_Users.OwnerId AND ADHomeDirectory = '$directory'"
	Write-Host "Query:" $query -ForegroundColor Green
	$report = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource

	# Parse Data
	$unixEpoch = Get-Date -Date "01/01/1970"
	for ($i = 0; $i -lt @($report).Count; $i++) {
		$report[$i].FileSizeDisk = [Math]::Round(($report[$i].FileSizeDisk / 1000), 2) # Convert to GB
		$report[$i].FileSizeCrawl = [Math]::Round(($report[$i].FileSizeCrawl / 1000), 2) # Convert to GB
		$report[$i].CreatedDate = $unixEpoch.AddSeconds($report[$i].createdDate)
		if ($report[$i].Extensions.Length -gt 32767) { # Check if Extensions field is too long
			Write-Host "Extensions field too long!" -ForegroundColor Red
			$report[$i].Extensions = "Too long!  See database"
		}
	}
	
	$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
	$logFile = $PSScriptRoot + "\report_$timestamp.xlsx"

	$report | Export-Excel $logFile -WorksheetName "Report" -Title "Report"	-TitleSize 18 -TitleBold -AutoSize -MaxAutoSizeRows 2
	
	$excel = Open-ExcelPackage -Path $logFile
	$sheet = $excel.Workbook.Worksheets[1]
	Set-ExcelRange -Range $sheet.Cells["H2:H2"] -Value "FileSizeDisk (GB)" -AutoSize
	Set-ExcelRange -Range $sheet.Cells["I2:I2"] -Value "FileSizeCrawl (GB)" -AutoSize
	Set-Column -Worksheet $sheet -Column 1 -Width 3
	Set-Column -Worksheet $sheet -Column 15 -Width 15
	Close-ExcelPackage $excel

	<#
	Write-Host "running query"
	#Write-Host $users
	$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
	$logFile = $PSScriptRoot + "\report_" + $timestamp + "$ownerId.csv"
	$report  | Export-Csv $logFile
	

	# Microsoft errors query
	$query = "	SELECT OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
				FROM Files_Batch_Users, Files_OneDrive
				WHERE ADHomeDirectory = '$directory' AND Files_Batch_Users.id = Files_OneDrive.OwnerId AND Error <> '' AND Error <> ' '"
	Write-Host "Query:" $query -ForegroundColor Green
	$report = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	Write-Host "running query"
	$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
	$logFile = $PSScriptRoot + "\msft_errors_" + $timestamp + "$ownerId.csv"
	$report  | Export-Csv $logFile
	#>



	# Excel spreadsheet

	$query = "	SELECT OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
				FROM Files_Batch_Users, Files_OneDrive
				WHERE ADHomeDirectory = '$directory' AND Files_Batch_Users.id = Files_OneDrive.OwnerId AND Error <> '' AND Error <> ' '"
	Write-Host "Query:" $query -ForegroundColor Green
	$report = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	Write-Host "running query"

	$errors = $report | Group-Object -Property Error | Sort-Object -Property Count -Descending | Select-Object Count, Name
	$chartDef = New-ExcelChart -Title "Top 5 Errors" -ChartType ColumnClustered -XRange "Name" -YRange "Count" -Width 750 -YAxisTitleText "Count" `
				-NoLegend -Row 7 -Column 1
	$errors | Select-Object -First 5 | Export-Excel $logFile -WorksheetName "Error Report Overview" -Title "Error Report Overview" -TitleSize 18 `
		-TitleBold -ExcelChartDefinition $chartDef -AutoSize -AutoNameRange
	$errors | Export-Excel $logFile -WorksheetName "Full Errors Count" -Title "Full Errors List" -TitleSize 18 -TitleBold -AutoFilter
	$report | Export-Excel $logFile -WorksheetName "All Errors" -Title "Errors" -TitleSize 18 -TitleBold -AutoFilter -AutoSize -MaxAutoSizeRows 2

	<#
	$excel = Open-ExcelPackage -Path $logFile
	$sheet = $excel.Workbook.Worksheets[3]
	Add-ExcelChart -Worksheet $sheet -ChartType ColumnClustered -XRange "B3:B7" -YRange "A3:A7" -Width 750 -NoLegend -Column 3 -Row 2
	Close-ExcelPackage $excel
	#>
	
	<#
	#No Access Report
	$logFile = $PSScriptRoot + "\no_access_" + $timestamp + "$ownerId.xlsx"
	$report | Where-Object {$_.Error -eq "No Access"} | Group-Object -Property ParentFolder | Sort-Object -Propert Count -Descending | Select-Object Count, Name | 
		Export-Excel $logFile -WorksheetName "No Access" -Title "No Access" -TitleSize 18 -TitleBold
	#>

	Write-Host "OwnerID:   $ownerId"

}

CreateNewDirectoryEntry $startDirectory
InitPreMigrationMaster $startDirectory
GeneratePostScanReport $startDirectory
Pause

# OLD!! usage run as admin --> .\pre_migration_master.ps1 -startDirectory "c:\test"