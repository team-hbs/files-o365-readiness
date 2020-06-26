param (
	[string]$mode,
	[string]$source,
	[string]$report
)

# Display window to browse for folder
if (($mode -eq "single") -and ($source -eq "")) {
	Add-Type -AssemblyName System.Windows.Forms
	$FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
	$FileBrowser.ShowDialog()
	$source = $FileBrowser.SelectedPath
}

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

# Create the path to the crawl script and run it
$crawlPath = $PSScriptRoot + "\crawl_v10.ps1"
. $crawlPath
#. .\office_cleanup.ps1

function GetUsers($batchNumber) {
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

function GetNewBatch($directory) {
	$user = $null
	Write-Host $batchNumber
	$query = "	SELECT *
				FROM Files_Batch_Users
				ORDER BY Id Desc
				LIMIT 1"
	Write-Host "Query:" $query -ForegroundColor Green
	$user = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	Write-Host "running query"
	#Write-Host $users
	return $user
}

function CreateNewDirectoryEntry($directory) {
 
	$query = "Insert INTO Files_Batch_Users (SamAccountName, ADhomeDirectory,BatchNumber) VALUES ('jbaldwin', '$directory', 0)"
    #Write-Host "Query:" $query -ForegroundColor Green 	   
	Invoke-SqliteQuery -Query $Query -DataSource $global:DataSource
}

function InitPreMigrationMaster($directory) {
    #$users = GetUsers $batchNumber
	$user = GetNewBatch $directory
	Write-Host $user -ForegroundColor Cyan
	$path = $user.ADhomedirectory
	$ownerId = $user.Id
	
	#Write-Host "OwnerId:" $ownerId -ForegroundColor Yellow
	ClearCrawlData $ownerId
	$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
	$logFile = $PSScriptRoot + "\crawl_" + $timestamp + "$ownerId.csv"
	$errors = InitCrawl $ownerId $path $false |  Select-Object FileName,Message,ParentFolderCurrent, Query
	if ($global:currentErrorCount -gt 0) {
		$errors | Export-Csv $logFile
	}
}

function ClearCrawlData($ownerId) {
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

function OfficeConversionTest() {
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

function GeneratePostScanReport ($directorySource) {
	if ($mode -eq "import" -and $report -eq "overall") {
		foreach ($row in $directorySource) {
			$directory = $row.HomeDirectory
			Write-Host "Directory: $directory"
			$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
						FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate 
						FROM Files_Batch_Users,Files_Users 
						WHERE Files_Batch_Users.Id = Files_Users.OwnerId AND ADHomeDirectory = '$directory'"
			Write-Host "Query:" $query -ForegroundColor Green
			$queryReturn += @(Invoke-SqliteQuery -Query $query -DataSource $global:DataSource)
			Write-Host "running query"
		}

		$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
		$logFile = $PSScriptRoot + "\report_$timestamp.xlsx"

	} else {
		$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
					FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate 
					FROM Files_Batch_Users,Files_Users 
					WHERE Files_Batch_Users.Id = Files_Users.OwnerId AND ADHomeDirectory = '$directorySource'"
		Write-Host "Query:" $query -ForegroundColor Green
		$queryReturn = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
		Write-Host "running query"

		$timestamp =  Get-Date -f _MM_dd_HH_mm_ss

		$logFile = $PSScriptRoot + "\report_$timestamp" + ($queryReturn.Id | Select-Object -Last 1) + ".xlsx"
	}

	GenerateXlsxReportMain $logFile $queryReturn

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
	$queryReturn = $null
	if ($mode -eq "import" -and $report -eq "overall") {
		foreach ($row in $directorySource) {
			$directory = $row.HomeDirectory
			Write-Host "Directory Errors: $directory"
			$query = "	SELECT OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
						FROM Files_Batch_Users, Files_OneDrive
						WHERE ADHomeDirectory = '$directory' AND Files_Batch_Users.id = Files_OneDrive.OwnerId AND Error <> '' AND Error <> ' '"
			Write-Host "Query:" $query -ForegroundColor Green
			$queryReturn += @(Invoke-SqliteQuery -Query $query -DataSource $global:DataSource)
		}
	} else {
		$query = "	SELECT OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
					FROM Files_Batch_Users, Files_OneDrive
					WHERE ADHomeDirectory = '$directorySource' AND Files_Batch_Users.Id = Files_OneDrive.OwnerId AND Error <> '' AND Error <> ' '"
		Write-Host "Query:" $query -ForegroundColor Green
		$queryReturn = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
		Write-Host "running query"
	}

	GenerateXlsxReportErrors $logFile $queryReturn
	
	<#
	#No Access Report
	$logFile = $PSScriptRoot + "\no_access_" + $timestamp + "$ownerId.xlsx"
	$report | Where-Object {$_.Error -eq "No Access"} | Group-Object -Property ParentFolder | Sort-Object -Propert Count -Descending | Select-Object Count, Name | 
		Export-Excel $logFile -WorksheetName "No Access" -Title "No Access" -TitleSize 18 -TitleBold
	#>

}

function GenerateOverallReport {
	$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
					FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate 
					FROM Files_Batch_Users,Files_Users
					WHERE Files_Batch_Users.Id = Files_Users.OwnerId"
	Write-Host "Query:" $query -ForegroundColor Green
	$queryReturn = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	Write-Host "running query"

	$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
	$logFile = $PSScriptRoot + "\report_$timestamp.xlsx"
	GenerateXlsxReportMain $logFile $queryReturn

	$query = "	SELECT OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
				FROM Files_Batch_Users, Files_OneDrive
				WHERE Files_Batch_Users.Id = Files_OneDrive.OwnerId AND Error <> '' AND Error <> ' '"
	Write-Host "Query:" $query -ForegroundColor Green
	$queryReturn = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	Write-Host "running query"
	
	GenerateXlsxReportErrors $logFile $queryReturn
}

function GenerateSingleReports {
	$query = "	SELECT Id FROM Files_Batch_Users"
	Write-Host "Query:" $query -ForegroundColor Green
	$ids = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource

	foreach ($id in $ids.id) {
		$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
					FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate 
					FROM Files_Batch_Users,Files_Users
					WHERE Files_Batch_Users.Id = $id AND Files_Batch_Users.Id = Files_Users.OwnerId"
		Write-Host "Query:" $query -ForegroundColor Green
		$queryReturn = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
		Write-Host "running query"

		$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
		$logFile = $PSScriptRoot + "\report_$timestamp" + "_$id" + ".xlsx"
		GenerateXlsxReportMain $logFile $queryReturn

		$query = "	SELECT OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
					FROM Files_Batch_Users, Files_OneDrive
					WHERE Files_Batch_Users.Id = $id AND Files_Batch_Users.Id = Files_OneDrive.OwnerId AND Error <> '' AND Error <> ' '"
		Write-Host "Query:" $query -ForegroundColor Green
		$queryReturn = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
		Write-Host "running query"
	
		GenerateXlsxReportErrors $logFile $queryReturn
	}
}

function GenerateXlsxReportMain ($logFile, $reportData) {
	# Parse Data
	$unixEpoch = Get-Date -Date "01/01/1970"
	for ($i = 0; $i -lt @($reportData).Count; $i++) {
		$reportData[$i].FileSizeDisk = [Math]::Round(($reportData[$i].FileSizeDisk / 1000), 2) # Convert to GB
		$reportData[$i].FileSizeCrawl = [Math]::Round(($reportData[$i].FileSizeCrawl / 1000), 2) # Convert to GB
		$reportData[$i].CreatedDate = $unixEpoch.AddSeconds($reportData[$i].createdDate)
		if ($reportData[$i].Extensions.Length -gt 32767) { # Check if Extensions field is too long
			Write-Host "Extensions field too long!" -ForegroundColor Red
			$reportData[$i].Extensions = "Too long!  See database"
		}
	}
	
	$reportData | Export-Excel $logFile -WorksheetName "Report" -Title "Report"	-TitleSize 18 -TitleBold -AutoSize -MaxAutoSizeRows 2
	
	$excel = Open-ExcelPackage -Path $logFile
	$sheet = $excel.Workbook.Worksheets["Report"]
	Set-ExcelRange -Range $sheet.Cells["H2:H2"] -Value "FileSizeDisk (GB)" -AutoSize
	Set-ExcelRange -Range $sheet.Cells["I2:I2"] -Value "FileSizeCrawl (GB)" -AutoSize
	Set-Column -Worksheet $sheet -Column 1 -Width 3
	Set-Column -Worksheet $sheet -Column 15 -Width 15
	Close-ExcelPackage $excel
}

function GenerateXlsxReportErrors ($logFile, $reportData) {
	if ($reportData.Count -gt 0) {
		$errors = $reportData | Group-Object -Property Error | Sort-Object -Property Count -Descending | Select-Object Count, Name
		$errors | Export-Excel $logFile -WorksheetName "Error Report" -Title "Error Report" -TitleSize 18 -TitleBold -AutoFilter
		$reportData | Export-Excel $logFile -WorksheetName "All Errors" -Title "Errors" -TitleSize 18 -TitleBold -AutoFilter -AutoSize -MaxAutoSizeRows 2

		
		$excel = Open-ExcelPackage -Path $logFile
		$sheet = $excel.Workbook.Worksheets["Error Report"]
		Set-Column -Worksheet $sheet -Column 2 -Width 65
		Set-Column -Worksheet $sheet -Column 3 -Value " "
		Add-ExcelChart -Worksheet $sheet -ChartType ColumnClustered -Title "Top 5 Errors" -XRange "B3:B7" -YRange "A3:A7" -Width 650 -YAxisTitleText "Count" -NoLegend -Column 3 -Row 2
		Close-ExcelPackage $excel
	} else {
		"NO ERRORS" | Export-Excel $logFile -WorksheetName "Error Report" -Title "Error Report" -TitleSize 18 -TitleBold
	}
}

if ($mode -eq "single") {
	CreateNewDirectoryEntry $source
	InitPreMigrationMaster $source
	if ($report -ne "") {
		GeneratePostScanReport $source
	}
} elseif ($mode -eq "import") {
	$rows = Import-Csv $source
	$directoriesCount = ($rows | Measure-Object).Count
	$currentDirectory = 0
	foreach ($row in $rows) {
		Write-Progress -Id 2 -Activity "Directories" -Status "Progress: $currentDirectory / $directoriesCount Directories" -PercentComplete ($currentDirectory / $directoriesCount * 100)
		$currentDirectory++
		Write-Host "source:  $directory"
		$directory = $row.HomeDirectory
        if ($directory.Trim() -ne "")
        {
		    CreateNewDirectoryEntry $directory
		    InitPreMigrationMaster $directory
		    if ($report -eq "single") {
			    GeneratePostScanReport $directory
		    }
        }
	}
	if ($report -eq "overall") {
		GeneratePostScanReport $rows
	}
} elseif ($mode -eq "report") {
	if ($report -eq "single") {
		GenerateSingleReports
	} elseif ($report -eq "overall") {
		GenerateOverallReport
	}
} else {
	Write-Host "Please select a valid mode!"
}


Pause

# OLD!! usage run as admin --> .\pre_migration_master.ps1 -startDirectory "c:\test"