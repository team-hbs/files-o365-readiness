param (
	[string]$connectionString,
	[string]$mode = $null,
	[string]$path = $null,
	[string]$report,
	[string]$database,
	[string]$configMode,
	[string]$notifications,
	[string]$email,
	[string]$key = '',
	[string]$value = '',
    [int]$BatchNumber = -1,
    [int]$SourceId = -1,
	[boolean]$encrypt = $false,
	[string] $lastModifiedDate = $null
)

if ($email -eq $null)
{
	$email = ''
}

$OwnerId = $SourceId

if ($PSVersionTable.PSVersion.Major -ge 5)
{
   
}
else
{
	write-host 'Powershell Version 5 or greater is required to run this script' -f Yellow
	break
}


if ($mode -ne 'Install')
{
	if ((Get-Module -ListAvailable -Name PSSQLite) -ne $null) 
	{
		Import-Module -Name PSSQLite
	} 
	else 
	{
		write-host "Missing PSSQLite Please run -mode 'Install'" -f Yellow
		break
	}
	# Module for interacting with xlsx files
	if ((Get-Module -ListAvailable -Name ImportExcel) -ne $null) 
	{
		Import-Module -Name ImportExcel
	} 
	else 
	{
		write-host "Missing ImportExcel Please run -mode 'Install'"
		$okToRun = $false
		break
	}
}


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


$smtp = "smtp.gmail.com"
$from = "heartlandpowershellscripts@gmail.com"
$username = "heartlandpowershellscripts@gmail.com"
$password = ConvertTo-SecureString -String "heartland123" -AsPlainText -Force
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password


  

function SqlQueryInsert($query) {
    write-host $query -f Yellow
	$null = @(
		if ($global:SqlServer) {
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

function AddEvent($ownerId, $eventType)
{
	$eventDate = Get-Date
	$query = "INSERT INTO Event (OwnerId, EventDate, EventType) VALUES($ownerId, '$eventDate', '$eventType')"
	SqlQueryInsert $query
}


function sendNotification($message) {
	if ($notifications -eq "on") {
		$subject = "Powershell Crawl"
		$to = $email.Split(";")
		Send-MailMessage -To $to -From $from -Subject $subject -Body $message -SmtpServer $smtp -Credential $credential -UseSsl -Port 587 -DeliveryNotificationOption Never
	}
}

function GetUsers($batchNumber) {
	$users = $null
	Write-Host $batchNumber
	$query = "	SELECT Source.ADhomeDirectory, Source.Id 
				FROM Source 
				WHERE Source.BatchNumber = '$batchNumber'"
	Write-Host "Query:" $query -ForegroundColor Green
	$users = SqlQueryReturn($query)
	Write-Host "running query"
	#Write-Host $users
	return $users
}

function GetOldOfficeDocuments($directory)
{
	$documents = $null
	$query = "	SELECT ScanFile.Path as Path, ScanFile.Id as Id, ScanFile.OwnerId as OwnerId 
				FROM ScanFile, Source 
				WHERE (Extension = 'xls' OR Extension = 'doc' OR Extension = 'ppt') 
				AND Source.Id = ScanFile.OwnerID AND Source.ADhomeDirectory = '$directory'"
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
	if ($global:SqlServer) {
		$query = "	SELECT TOP 1 *
					FROM Source
					ORDER BY Id Desc"
	} else {
		$query = "	SELECT *
					FROM Source
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
 
	$query = "Insert INTO Source (SamAccountName, ADhomeDirectory,BatchNumber) VALUES ('jbaldwin', '$directory', 0)"
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
	
	$noOfficeValue = [String] (GetConfig('NoOffice') )
    if ($noOfficeValue -eq 'true')
    {
		$noOffice = $true
	}
	else
	{
		$noOffice = $false
	}
	
	$errors = InitCrawl $ownerId $path $false $noOffice |  Select-Object FileName,Message,ParentFolderCurrent, Query
	if ($global:currentErrorCount -gt 0) {
		$errors | Export-Csv $logFile
	}
}

function ClearCrawlData($ownerId) {
    try
    {
		$query = "DELETE FROM ScanJob WHERE OwnerId = '$ownerId'"
		SqlQueryInsert($query)
		$query = "DELETE FROM ScanFile WHERE OwnerId = '$ownerId'"
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
			if ($global:SqlServer) {
				$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
							FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate, CONVERT(DATETIME, '2001-01-01', 102) as DateCreated
							FROM Source,ScanJob 
							WHERE Source.Id = ScanJob.OwnerId AND ADHomeDirectory = '$directory'"
			} else {
				$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
							FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate 
							FROM Source,ScanJob 
							WHERE Source.Id = ScanJob.OwnerId AND ADHomeDirectory = '$directory'"
			}
			Write-Host "Query:" $query -ForegroundColor Green
			$queryReturn += @(SqlQueryReturn($query))
			Write-Host "running query"
		}

		$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
		$logFile = $PSScriptRoot + "\report_$timestamp.xlsx"

	} else {
		if ($global:SqlServer) {
			$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
						FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate, CONVERT(DATETIME, '2001-01-01', 102) as DateCreated
						FROM Source,ScanJob 
						WHERE Files_Batch_Users.Id = Files_Users.OwnerId AND ADHomeDirectory = '$directorySource'"
		} else {
			$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
						FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate 
						FROM Source,ScanJob 
						WHERE Source.Id = ScanJob.OwnerId AND ADHomeDirectory = '$directorySource'"
		}
		Write-Host "Query:" $query -ForegroundColor Green
		$queryReturn = SqlQueryReturn($query)
		Write-Host "running query"

		$timestamp =  Get-Date -f _MM_dd_HH_mm_ss

		$logFile = $PSScriptRoot + "\report_$timestamp" + ($queryReturn.Id | Select-Object -Last 1) + ".xlsx"
	}

	#Write-Host "queryReturn"

	GenerateXlsxReportMain $logFile $queryReturn

	# Excel spreadsheet
	$queryReturn = $null
	if ($mode -eq "import" -and $report -eq "overall") {
		foreach ($row in $directorySource) {
			$directory = $row.HomeDirectory
			Write-Host "Directory Errors: $directory"
			$query = "	SELECT OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
						FROM Source, ScanFile
						WHERE ADHomeDirectory = '$directory' AND Source.id = ScanFile.OwnerId AND Error <> '' AND Error <> ' '"
			Write-Host "Query:" $query -ForegroundColor Green
			$queryReturn += @(SqlQueryReturn($query))
		}
	} else {
		$query = "	SELECT OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
					FROM Source, ScanFile
					WHERE ADHomeDirectory = '$directorySource' AND Source.Id = ScanFile.OwnerId AND Error <> '' AND Error <> ' '"
		Write-Host "Query:" $query -ForegroundColor Green
		$queryReturn = SqlQueryReturn($query)
		Write-Host "running query"
	}

	GenerateXlsxReportErrors $logFile $queryReturn

}

function GenerateOverallReport {
	if ($global:SqlServer) {
		$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
						FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate, CONVERT(DATETIME, '2001-01-01', 102) as DateCreated
						FROM Source,ScanJob
						WHERE Source.Id = ScanJob.OwnerId"
	} else {
		$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
						FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate 
						FROM Source,ScanJob
						WHERE Source.Id = ScanJob.OwnerId"
	}
	Write-Host "Query:" $query -ForegroundColor Green
	$queryReturn = SqlQueryReturn($query)
	Write-Host "running query"

	$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
	$logFile = $PSScriptRoot + "\report_$timestamp.xlsx"
	GenerateXlsxReportMain $logFile $queryReturn

	$query = "	SELECT OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
				FROM Source, ScanFile
				WHERE Source.Id = ScanFile.OwnerId AND Error <> '' AND Error <> ' '"
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
		if ($global:SqlServer) {
			$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
						FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate, CONVERT(DATETIME, '2001-01-01', 102) as DateCreated
						FROM Source,ScanJob
						WHERE Source.Id = $id AND Source.Id = ScanJob.OwnerId"
		} else {
			$query = "	SELECT Id, SamAccountName, ADHomeDirectory, FileCountDisk, FileCountCrawl, MacroCount, Extensions,
						FileSizeDisk, FileSizeCrawl, ErrorCount, OfficeErrorCount, OldOfficeCount, PathLengthCount, NoAccessCount, CreatedDate 
						FROM Source,ScanJob
						WHERE Source.Id = $id AND Source.Id = ScanJob.OwnerId"
		}
		Write-Host "Query:" $query -ForegroundColor Green
		$queryReturn = SqlQueryReturn($query)
		Write-Host "running query"

		$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
		$logFile = $PSScriptRoot + "\report_$timestamp" + "_$id" + ".xlsx"
		GenerateXlsxReportMain $logFile $queryReturn

		$query = "	SELECT OwnerId, SamAccountName, BatchNumber, ADHomeDirectory, FileName, Extension, Path, ParentFolder, Error
					FROM Source, ScanFile
					WHERE Source.Id = $id AND Source.Id = ScanFile.OwnerId AND Error <> '' AND Error <> ' '"
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

			if ($global:SqlServer) {
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

			if ($global:SqlServer) {
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

	if ($global:SqlServer) {
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

function GetImportFile($startsIn)
{  
    $fileName = $null
    $null = @(
        [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) |
        Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $startsIn
        $OpenFileDialog.filter = 'All files (*.xlsx)| *.xlsx'
        $OpenFileDialog.ShowDialog() | Out-Null
        $fileName = $OpenFileDialog.filename
    )
    return $fileName
} 


# Create the path to the crawl script and run it
$crawlPath = $PSScriptRoot + "\crawl_v11.ps1"
. $crawlPath


if ($mode -eq "single") 
{
	$crawlMonitor = $null
	if ($notifications -eq "on") {
		if ($connectionString -eq "") {
			$crawlMonitor = Start-Process PowerShell.exe -PassThru -WindowStyle Hidden -Argument "-NoExit -NoProfile -ExecutionPolicy Bypass -File .\crawlmonitor.ps1 -email ""$email"""
		} else {
			$crawlMonitor = Start-Process PowerShell.exe -PassThru -WindowStyle Hidden -Argument "-NoExit -NoProfile -ExecutionPolicy Bypass -File .\crawlmonitor.ps1 -connectionString ""$connectionString"" -email ""$email"""
		}
	}
	CreateNewDirectoryEntry $path
	InitPreMigrationMaster $path
	if ($report -ne "") {
		GeneratePostScanReport $path
	}
	if ($notifications -eq "on") {
		Stop-Process $crawlMonitor
	}
	sendNotification("Scan complete!")

	$computerName = [System.Net.Dns]::GetHostName()
	$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
	$makeLocation = "https://migrationstoragehbstemp.blob.core.windows.net/scans/" + $computerName + "?sv=2020-02-10&ss=bfqt&srt=sco&sp=rwdlacupx&se=2022-03-27T04:23:01Z&st=2021-03-26T20:23:01Z&spr=https&sig=xsQFrHErUapJLzQzFQ7w%2BTjyARMo5vXE1iYgr01ZcDU%3D"
	$uploadLocation = "https://migrationstoragehbstemp.blob.core.windows.net/scans/" + $computerName + "/" + $timestamp + "/?sv=2020-02-10&ss=bfqt&srt=sco&sp=rwdlacupx&se=2022-03-27T04:23:01Z&st=2021-03-26T20:23:01Z&spr=https&sig=xsQFrHErUapJLzQzFQ7w%2BTjyARMo5vXE1iYgr01ZcDU%3D"
	
	$makeString = "azcopy make ""$makeLocation"""
	Write-Host $makeString
	#Invoke-Expression $makeString
	
	$uploadString = "azcopy copy ""$PSScriptRoot\FilesToO365.db"" ""$uploadLocation"""
	Write-Host $uploadString
	#Invoke-Expression $uploadString
}
elseif($mode -eq 'ConfigReport')
{
        $query = 'SELECT * 
                FROM Config
                ORDER BY Key ASC
          '
        $configs =  Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
        $configs | Out-GridView
}
elseif($mode -eq 'BatchReport')
{
    $query = "SELECT Source.Id as OwnerId, BatchNumber, ADHomeDirectory as SourceDirectory, FileSizeDisk, OldOfficeCount, PathLengthCount,NoAccessCount, CreatedDate
        FROM Source
        LEFT JOIN ScanJob
        ON Source.Id = ScanJob.OwnerId"
    $batches = SqlQueryReturn($query)
    $batches | Out-GridView
}
elseif($mode -eq 'Delete')
{

}
elseif($mode -eq 'Clear')
{


}
elseif($mode -eq 'CleanUp')
{
    if ($batchNumber -ne -1)
    {
        $query = "SELECT Source.ADhomeDirectory, Source.Id 
				FROM Source 
				WHERE Source.BatchNumber = $batchNumber
                ORDER BY Id DESC"
	    Write-Host "Query:" $query -ForegroundColor Green
        $sources = SqlQueryReturn($query)
       
        foreach($row in $sources)
        {
			Invoke-Expression ".\office_cleanup.ps1	-ownerId $ownerId -doConvert $false" 
        }   
    }
    elseif($ownerId -ne -1)
    {
        $query = "SELECT Source.ADhomeDirectory, Source.Id 
				FROM Source 
				WHERE Source.Id = $ownerId"
	    Write-Host "Query:" $query -ForegroundColor Green
	    $sources = SqlQueryReturn($query)
        foreach($row in $sources)
        {
			Invoke-Expression ".\office_cleanup.ps1	-ownerId $ownerId -doConvert $false" 
        }
    }
}
elseif ($mode -eq "Scan") 
{
    if ($batchNumber -ne -1)
    {
        $query = "SELECT Source.ADhomeDirectory, Source.Id 
				FROM Source 
				WHERE Source.BatchNumber = $batchNumber
                ORDER BY Id DESC"
	    Write-Host "Query:" $query -ForegroundColor Green
        $source = SqlQueryReturn($query)
        $directoriesCount = ($source | Measure-Object).Count
	    $currentDirectory = 0
		
		$noOfficeValue = [String] (GetConfig('NoOffice') )
		if ($noOfficeValue -eq 'true')
		{
			$noOffice = $true
		}
		else
		{
			$noOffice = $false
		}
		
        foreach($row in $source)
        {
            Write-Progress -Id 2 -Activity "Directories" -Status "Progress: $currentDirectory / $directoriesCount Directories" -PercentComplete ($currentDirectory / $directoriesCount * 100)
	    	$currentDirectory++
            $path = $row.ADHomeDirectory
            $ownerId = $row.Id
			#$ownerId, $startPath, $doConvert, $noOffice)
			if ($lastModifiedDate -eq $null)
			{
				InitCrawl -ownerId $ownerId -startPath $path -doConvert $false -noOffice $noOffice
			}
			else
			{
				InitCrawl -ownerId $ownerId -startPath $path -doConvert $false -noOffice $noOffice -lastModifiedDate $lastModifiedDate
			}
            $currentDirectory++
        }   
    }
    elseif($ownerId -ne -1)
    {
        $query = "SELECT Source.ADhomeDirectory, Source.Id 
				FROM Source 
				WHERE Source.Id = $ownerId"
	    Write-Host "Query:" $query -ForegroundColor Green
	    $source = SqlQueryReturn($query)
        $directoriesCount = ($source | Measure-Object).Count
	    $currentDirectory = 0
		$noOfficeValue = [String] (GetConfig('NoOffice') )
		if ($noOfficeValue -eq 'true')
		{
			$noOffice = $true
		}
		else
		{
			$noOffice = $false
		}
        foreach($row in $source)
        {
            Write-Progress -Id 2 -Activity "Directories" -Status "Progress: $currentDirectory / $directoriesCount Directories" -PercentComplete ($currentDirectory / $directoriesCount * 100)
	    	$currentDirectory++
            $path = $row.ADHomeDirectory
			if ($lastModifiedDate -eq $null)
			{
				InitCrawl -ownerId $ownerId -startPath $path -doConvert $false -noOffice $noOffice
			}
			else
			{
				InitCrawl -ownerId $ownerId -startPath $path -doConvert $false -noOffice $noOffice -lastModifiedDate $lastModifiedDate
			}
            $currentDirectory++
        }
    }
    else
    {
	    $query = "SELECT Source.ADhomeDirectory, Source.Id 
				FROM Source 
				ORDER BY Id
				"
	    Write-Host "Query:" $query -ForegroundColor Green
	    $source = SqlQueryReturn($query)
        $directoriesCount = ($source | Measure-Object).Count
	    $currentDirectory = 0
		$noOfficeValue = [String] (GetConfig('NoOffice') )
		if ($noOfficeValue -eq 'true')
		{
			$noOffice = $true
		}
		else
		{
			$noOffice = $false
		}
        foreach($row in $source)
        {
            Write-Progress -Id 2 -Activity "Directories" -Status "Progress: $currentDirectory / $directoriesCount Directories" -PercentComplete ($currentDirectory / $directoriesCount * 100)
	    	$currentDirectory++
            $path = $row.ADHomeDirectory
			$ownerId = $row.Id
            InitCrawl $ownerId $path $false $noOffice
            $currentDirectory++
        }
    }

	$computerName = [System.Net.Dns]::GetHostName()
	$timestamp =  Get-Date -f _MM_dd_HH_mm_ss
	$makeLocation = "https://migrationstoragehbstemp.blob.core.windows.net/scans/" + $computerName + "?sv=2020-02-10&ss=bfqt&srt=sco&sp=rwdlacupx&se=2022-03-27T04:23:01Z&st=2021-03-26T20:23:01Z&spr=https&sig=xsQFrHErUapJLzQzFQ7w%2BTjyARMo5vXE1iYgr01ZcDU%3D"
	$uploadLocation = "https://migrationstoragehbstemp.blob.core.windows.net/scans/" + $computerName + "/" + $timestamp + "/?sv=2020-02-10&ss=bfqt&srt=sco&sp=rwdlacupx&se=2022-03-27T04:23:01Z&st=2021-03-26T20:23:01Z&spr=https&sig=xsQFrHErUapJLzQzFQ7w%2BTjyARMo5vXE1iYgr01ZcDU%3D"
	
	$makeString = "azcopy make ""$makeLocation"""
	Write-Host $makeString
	#Invoke-Expression $makeString
	
	$uploadString = "azcopy copy ""$PSScriptRoot\FilesToO365.db"" ""$uploadLocation"""
	Write-Host $uploadString
	#Invoke-Expression $uploadString
}
elseif($mode -eq 'Install')
{
	unblock-file -path .\pre_migration_master.ps1
	unblock-file -path .\crawl_v11.ps1
	unblock-file -path .\create_db.ps1
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	Install-Module -Name PSSQLite
    Install-Module -Name ImportExcel
	Invoke-Expression "$PSScriptRoot\installAzCopy -InstallPath ""$PSScriptRoot"""
}
elseif ($mode -eq 'Import')
{
	$query = "SELECT * 
		FROM Source"
	$dbSources = SqlQueryReturn($query)

	$path = GetImportFile (Get-Location)
	$sources = Import-Excel $path -WorkSheetname 'Source'
	$counter = 0
	foreach($source in $sources)
	{
			try
			{
				$batchNumber = $source.BatchNumber
				$adHomeDirectory = $source.SourceDirectory.Trim().ToLower()

				if ($adHomeDirectory -eq $null)
				{
					$adHomeDirectory = ''
				}
				if ($batchNumber -ne $null -AND $adHomeDirectory.Trim() -ne '')
				{
					$dbSourceId = $null
					foreach($dbSource in $dbSources)
					{
						if ($dbSource.ADHomeDirectory -eq $adHomeDirectory.ToLower().Trim())
						{
							$dbSourceId = $dbSource.Id
							break
						}
					}
					if ($dbSourceId -eq $null)
					{
						$query = "INSERT INTO Source (ADHomeDirectory, BatchNumber) VALUES ('$adHomeDirectory', $batchNumber)"
					}
					else
					{
						$query = "UPDATE Source SET BatchNumber = $batchNumber WHERE Id = $dbSourceId"
					}
					write-host $query
					write-host "Query:" $query -ForegroundColor Green 	   
					SqlQueryInsert($query)
				}
				
			}
			catch
			{
				write-host ''
			}
	}
}
elseif ($mode -eq 'SetConfig')
{
	SetConfig $key $value $encrypt
    write-host 'Updating' $key 'to' $value -f Green
}
elseif($mode -eq 'GetConfig')
{
	$configValue = GetConfig $key
    write-host 'Getting' $key 'Value:' $configValue -f Green
}
elseif($mode -eq 'CreateDatabase')
{
	if (GetConfig 'DatabaseMode' -eq 'SQLServer')
	{
		$create_db = $PSScriptRoot + "\create_db.ps1"
		. $create_db
	}
	else
	{
		write-host 'Cannot Create Database in SQLite Mode' -f Yellow
	}
}
elseif ($mode -eq "ClearDatabase") 
{
	if ($global:SqlServer) {
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

	$query = "SELECT * FROM Source"
	$Files_Batch_Users = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	if ($Files_Batch_Users -ne $null) {
		$query = "DELETE FROM Source"
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	}

	$query = "SELECT * FROM ScanFile"
	$Files_OneDrive = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	if ($Files_OneDrive -ne $null) {
		$query = "DELETE FROM Files_OneDrive"
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	}

	$query = "SELECT * FROM ScanJob"
	$Files_Users = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	if ($Files_Users -ne $null) {
		$query = "DELETE FROM ScanJob"
		Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	}

	$query = "	UPDATE sqlite_sequence SET seq = 0 WHERE name = 'Config';
				UPDATE sqlite_sequence SET seq = 0 WHERE name = 'Source';
				UPDATE sqlite_sequence SET seq = 0 WHERE name = 'ScanFile';
				UPDATE sqlite_sequence SET seq = 0 WHERE name = 'ScanJob';"
	Invoke-SqliteQuery -Query $query -DataSource $global:DataSource

	Write-Host "Cleared database!"
}
elseif($mode -eq $null)
{
    if ($batchNumber -ne -1)
    {

    }
} 
else 
{
	Write-Host "Please select a valid mode!"
}

