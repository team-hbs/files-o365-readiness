param (
	[string]$connectionString,
	[string]$mode,
	[string]$source,
	[string]$report,
	[string]$database,
	[string]$configMode,
	[string]$notifications,
	[string]$email
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
$global:SqlSever = $false


# Check config for SQL Server connection if none provided
if ($connectionString -eq "") {
	$query = "SELECT * FROM Config WHERE Id = 1"
	$SQLiteConfig = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	if ($SQLiteConfig -ne $null) {
		if ($SQLiteConfig.database -eq "sql-server") {
			if ($SQLiteConfig.connection -ne "") {
				$connectionString = $SQLiteConfig.connection
			}
		}
	}
}

# If using SQL Server connect
if ($connectionString -ne "") {
	$global:SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$global:SqlConnection.ConnectionString = $connectionString
	$global:SqlSever = $true
}

function SqlQueryInsert($query) {
	$null = @(
		if ($global:SqlSever) {
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd.CommandText = $query
			$SqlCmd.Connection = $global:SqlConnection
			if ($global:SqlConnection.State -ne 1)
			{
				$global:SqlConnection.Open()
			}
			$SqlCmd.ExecuteNonQuery()
			$global:SqlConnection.Close()

		} else {
			Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
		}
	)
}

function SqlQueryReturn($query) {
	if ($global:SqlSever) {
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


$query = "SELECT * FROM Config WHERE Id = 1"
$config = SqlQueryReturn($query)
if ($config -ne $null) {
	if (($mode -eq "") -and ($config.mode -ne "")) {
		$mode = $config.mode
	}
	if (($source -eq "") -and ($config.source -ne "")) {
		$source = $config.source
	}
	if (($report -eq "") -and ($config.report -ne "")) {
		$report = $config.report
	}
	if (($notifications -eq "") -and ($config.notifications -ne "")) {
		$notifications = $config.notifications
	}
	if (($email -eq "") -and ($config.email -ne "")) {
		$email = $config.email
	}
}

# Display window to browse for folder
if (($mode -eq "single") -and ($source -eq "")) {
	Add-Type -AssemblyName System.Windows.Forms
	$FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
	$FileBrowser.ShowDialog()
	$source = $FileBrowser.SelectedPath
}

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
	$users = SqlQueryReturn($query)
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
	$documents = SqlQueryReturn($query)
	Write-Host "running query"
	#Write-Host $documents
	return $documents
}

function GetNewBatch($directory) {
	$user = $null
	Write-Host $batchNumber
	$query = ""
	if ($global:SqlSever) {
		$query = "	SELECT TOP 1 *
					FROM Files_Batch_Users
					ORDER BY Id Desc"
	} else {
		$query = "	SELECT *
					FROM Files_Batch_Users
					ORDER BY Id Desc
					LIMIT 1"
	}
	Write-Host "Query:" $query -ForegroundColor Green
	$user = SqlQueryReturn($query)
	Write-Host "running query"
	#Write-Host $users
	return $user
}

function CreateNewDirectoryEntry($directory) {
 
	$query = "Insert INTO Files_Batch_Users (SamAccountName, ADhomeDirectory,BatchNumber) VALUES ('jbaldwin', '$directory', 0)"
    #Write-Host "Query:" $query -ForegroundColor Green 	   
	SqlQueryInsert($query)
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
		SqlQueryInsert($query)
		$query = "DELETE FROM Files_OneDrive WHERE OwnerId = '$ownerId'"
		SqlQueryInsert($query)
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
			if ($global:SqlSever) {
				$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
							FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate, CONVERT(DATETIME, '2001-01-01', 102) as DateCreated
							FROM Files_Batch_Users,Files_Users 
							WHERE Files_Batch_Users.Id = Files_Users.OwnerId AND ADHomeDirectory = '$directory'"
			} else {
				$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
							FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate 
							FROM Files_Batch_Users,Files_Users 
							WHERE Files_Batch_Users.Id = Files_Users.OwnerId AND ADHomeDirectory = '$directory'"
			}
			Write-Host "Query:" $query -ForegroundColor Green
			$queryReturn += @(SqlQueryReturn($query))
			Write-Host "running query"
		}

		$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
		$logFile = $PSScriptRoot + "\report_$timestamp.xlsx"

	} else {
		if ($global:SqlSever) {
			$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
						FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate, CONVERT(DATETIME, '2001-01-01', 102) as DateCreated
						FROM Files_Batch_Users,Files_Users 
						WHERE Files_Batch_Users.Id = Files_Users.OwnerId AND ADHomeDirectory = '$directorySource'"
		} else {
			$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
						FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate 
						FROM Files_Batch_Users,Files_Users 
						WHERE Files_Batch_Users.Id = Files_Users.OwnerId AND ADHomeDirectory = '$directorySource'"
		}
		Write-Host "Query:" $query -ForegroundColor Green
		$queryReturn = SqlQueryReturn($query)
		Write-Host "running query"

		$timestamp =  Get-Date -f _MM_dd_HH_mm_ss

		$logFile = $PSScriptRoot + "\report_$timestamp" + ($queryReturn.Id | Select-Object -Last 1) + ".xlsx"
	}

	#Write-Host "queryReturn"
	#$queryReturn
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
			$queryReturn += @(SqlQueryReturn($query))
		}
	} else {
		$query = "	SELECT OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
					FROM Files_Batch_Users, Files_OneDrive
					WHERE ADHomeDirectory = '$directorySource' AND Files_Batch_Users.Id = Files_OneDrive.OwnerId AND Error <> '' AND Error <> ' '"
		Write-Host "Query:" $query -ForegroundColor Green
		$queryReturn = SqlQueryReturn($query)
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
	if ($global:SqlSever) {
		$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
						FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate, CONVERT(DATETIME, '2001-01-01', 102) as DateCreated
						FROM Files_Batch_Users,Files_Users
						WHERE Files_Batch_Users.Id = Files_Users.OwnerId"
	} else {
		$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
						FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate 
						FROM Files_Batch_Users,Files_Users
						WHERE Files_Batch_Users.Id = Files_Users.OwnerId"
	}
	Write-Host "Query:" $query -ForegroundColor Green
	$queryReturn = SqlQueryReturn($query)
	Write-Host "running query"

	$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
	$logFile = $PSScriptRoot + "\report_$timestamp.xlsx"
	GenerateXlsxReportMain $logFile $queryReturn

	$query = "	SELECT OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
				FROM Files_Batch_Users, Files_OneDrive
				WHERE Files_Batch_Users.Id = Files_OneDrive.OwnerId AND Error <> '' AND Error <> ' '"
	Write-Host "Query:" $query -ForegroundColor Green
	$queryReturn = SqlQueryReturn($query)
	Write-Host "running query"
	
	GenerateXlsxReportErrors $logFile $queryReturn
}

function GenerateSingleReports {
	$query = "	SELECT Id FROM Files_Batch_Users"
	Write-Host "Query:" $query -ForegroundColor Green
	$ids = SqlQueryReturn($query)

	foreach ($id in $ids.id) {
		if ($global:SqlSever) {
			$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
						FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate, CONVERT(DATETIME, '2001-01-01', 102) as DateCreated
						FROM Files_Batch_Users,Files_Users
						WHERE Files_Batch_Users.Id = $id AND Files_Batch_Users.Id = Files_Users.OwnerId"
		} else {
			$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
						FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate 
						FROM Files_Batch_Users,Files_Users
						WHERE Files_Batch_Users.Id = $id AND Files_Batch_Users.Id = Files_Users.OwnerId"
		}
		Write-Host "Query:" $query -ForegroundColor Green
		$queryReturn = SqlQueryReturn($query)
		Write-Host "running query"

		$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
		$logFile = $PSScriptRoot + "\report_$timestamp" + "_$id" + ".xlsx"
		GenerateXlsxReportMain $logFile $queryReturn

		$query = "	SELECT OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
					FROM Files_Batch_Users, Files_OneDrive
					WHERE Files_Batch_Users.Id = $id AND Files_Batch_Users.Id = Files_OneDrive.OwnerId AND Error <> '' AND Error <> ' '"
		Write-Host "Query:" $query -ForegroundColor Green
		$queryReturn = SqlQueryReturn($query)
		Write-Host "running query"
	
		GenerateXlsxReportErrors $logFile $queryReturn
	}
}

function GenerateXlsxReportMain ($logFile, $reportData) {
	
	# Parse Data
	$unixEpoch = Get-Date -Date "01/01/1970"
	if (@($reportData).Count -eq 1) {
		$reportData.FileSizeDisk = [Math]::Round(($reportData.FileSizeDisk / 1000), 2) # Convert to GB
			$reportData.FileSizeCrawl = [Math]::Round(($reportData.FileSizeCrawl / 1000), 2) # Convert to GB

			if ($global:SqlSever) {
				$reportData.DateCreated = $unixEpoch.AddSeconds($reportData.createdDate)
			} else {
				$reportData.CreatedDate = $unixEpoch.AddSeconds($reportData.createdDate)
			}

			if ($reportData.Extensions.Length -gt 32767) { # Check if Extensions field is too long
				Write-Host "Extensions field too long!" -ForegroundColor Red
				$reportData.Extensions = "Too long!  See database"
			}

	} else {
		for ($i = 0; $i -lt @($reportData).Count; $i++) {
			$reportData[$i].FileSizeDisk = [Math]::Round(($reportData[$i].FileSizeDisk / 1000), 2) # Convert to GB
			$reportData[$i].FileSizeCrawl = [Math]::Round(($reportData[$i].FileSizeCrawl / 1000), 2) # Convert to GB

			if ($global:SqlSever) {
				$reportData[$i].DateCreated = $unixEpoch.AddSeconds($reportData[$i].createdDate)
			} else {
				$reportData[$i].CreatedDate = $unixEpoch.AddSeconds($reportData[$i].createdDate)
			}

			if ($reportData[$i].Extensions.Length -gt 32767) { # Check if Extensions field is too long
				Write-Host "Extensions field too long!" -ForegroundColor Red
				$reportData[$i].Extensions = "Too long!  See database"
			}
			
		}
	}

	if ($global:SqlSever) {
		$reportData = $reportData | Select-Object -Property Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions, FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, DateCreated
	} else {
		$reportData = $reportData | Select-Object -Property Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions, FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate
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

	if ((-not($reportData -eq $null)) -and @($reportData).Count -gt 0) {
		$reportData = $reportData | Select-Object -Property OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
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
} elseif ($mode -eq "config") {

	$query = "SELECT * FROM Config WHERE Id = 1"
	$SQLiteConfig = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	if ($SQLiteConfig -eq $null) {
		$query = "INSERT INTO Config DEFAULT VALUES"
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	}

	$query = "SELECT * FROM Config WHERE Id = 1"
	$config = SqlQueryReturn($query)
	if ($config -eq $null) {
		$query = "INSERT INTO Config DEFAULT VALUES"
		SqlQueryInsert($query)
	}

	if ($connectionString -ne "") {
		$query = "UPDATE Config SET connection = '$connectionString' WHERE Id = 1"
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	}
	if ($configMode -ne "") {
		$query = "UPDATE Config SET mode = '$configMode' WHERE Id = 1"
		SqlQueryInsert($query)
	}
	if ($source -ne "") {
		$query = "UPDATE Config SET source = '$source' WHERE Id = 1"
		SqlQueryInsert($query)
	}
	if ($report -ne "") {
		$query = "UPDATE Config SET report = '$report' WHERE Id = 1"
		SqlQueryInsert($query)
	}
	if ($database -ne "") {
		$query = "UPDATE Config SET database = ""$database"" WHERE Id = 1"
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	}
	if ($notifications -ne "") {
		$query = "UPDATE Config SET notifications = '$notifications' WHERE Id = 1"
		SqlQueryInsert($query)
	}
	if ($email -ne "") {
		$query = "UPDATE Config SET email = '$email' WHERE Id = 1"
		SqlQueryInsert($query)
	}

} elseif ($mode -eq "clear-database") {
	if ($global:SqlSever) {
		$query = "	DELETE FROM Config;
					DELETE FROM Files_Batch_Users;
					DELETE FROM Files_OneDrive;
					DELETE FROM Files_Users;
					DBCC CHECKIDENT (Config, RESEED, 0);
					DBCC CHECKIDENT (Files_Batch_Users, RESEED, 0);
					DBCC CHECKIDENT (Files_OneDrive, RESEED, 0);"
		SqlQueryInsert($query)
	}


	$query = "SELECT * FROM Config"
	$Config = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	if ($Config -ne $null) {
		$query = "DELETE FROM Config"
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	}

	$query = "SELECT * FROM Files_Batch_Users"
	$Files_Batch_Users = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	if ($Files_Batch_Users -ne $null) {
		$query = "DELETE FROM Files_Batch_Users"
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	}

	$query = "SELECT * FROM Files_OneDrive"
	$Files_OneDrive = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	if ($Files_OneDrive -ne $null) {
		$query = "DELETE FROM Files_OneDrive"
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	}

	$query = "SELECT * FROM Files_Users"
	$Files_Users = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	if ($Files_Users -ne $null) {
		$query = "DELETE FROM Files_Users"
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	}

	$query = "	UPDATE sqlite_sequence SET seq = 0 WHERE name = 'Config';
				UPDATE sqlite_sequence SET seq = 0 WHERE name = 'Files_Batch_Users';
				UPDATE sqlite_sequence SET seq = 0 WHERE name = 'Files_OneDrive';
				UPDATE sqlite_sequence SET seq = 0 WHERE name = 'Files_Users';"
	Invoke-SqliteQuery -Query $query -DataSource $global:DataSource

	Write-Host "Cleared database!"

} else {
	Write-Host "Please select a valid mode!"
}


Pause

# OLD!! usage run as admin --> .\pre_migration_master.ps1 -startDirectory "c:\test"