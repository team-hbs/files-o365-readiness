#1.22

#TODO: ADD CHECK FOR REQUIRED MODULES IF NOT IN INSTALL MODE
if ((Get-Module -ListAvailable -Name PSSQLite) -ne $null) {
    Import-Module -Name PSSQLite
} 
else {
   	write-host "Please run -mode 'Install'"
}

# Module for interacting with xlsx files
if ((Get-Module -ListAvailable -Name ImportExcel) -ne $null) {
    Import-Module -Name ImportExcel
} 
else {
   	write-host "Please run -mode 'Install'"
}

$global:DataSource = $PSScriptRoot + "\FilesToO365.db"

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
        $encryptedPassword = ConvertTo-SecureString -String ([string] $value)
        $value = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($encryptedPassword))
	}
	return $value
}

function SetConfig($key, $value, $encrypted)
{
	$null = @(
		if ($encrypted)
		{
			$value = ConvertTo-SecureString -String ([string] $value) -AsPlainText -Force  | ConvertFrom-SecureString
			$encrypted = 1
		}
		else
		{
			$encrypted = 0
		}
	    $query = "UPDATE Config SET Value='" + $value + "',Encrypted=" + $encrypted + " WHERE Key = '" + $key + "'"
        write-host $query -f yellow
        Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
        if ((GetConfig 'DatabaseMode') -eq 'SQLServer')
        {
            #query if value exists
            $serverName = $env:COMPUTERNAME
            $query = "SELECT * FROM GlobalConfig WHERE Server = '" + $serverName + "' AND [Key] = '" + $key + "'"
            $rows =  SqlQueryReturn -Query $query -DataSource $global:DataSource
            if ($rows -ne $null)
            {
               $query = "UPDATE GlobalConfig SET Value = '" + $value + "' WHERE Server = '" + $serverName + "' AND [Key] = '" + $key + "'"
               SqlQueryInsert -Query $query -DataSource $global:DataSource
            }
            else
            {
                $query = "INSERT INTO GlobalConfig ([Key],Server,Value) VALUES('" + $key + "','" + $serverName + "','" + $value + "')"
                SqlQueryInsert -Query $query -DataSource $global:DataSource
            }
        }
	)
}


function SqlQueryInsert($query) {
	$null = @(
		if ($global:SqlServer) {
            #write-host $query -f Yellow
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd.CommandText = $query
            $SqlCmd.CommandTimeout = 60
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
       write-host $query -f Cyan
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd.CommandText = $query
            $SqlCmd.CommandTimeout = 60
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


function OfficeMonitor()
{
    #start office monitor if it is not already running
    $noOfficeConfig = GetConfig 'NoOffice'
    if ($noOfficeConfig -eq 'false')
    {
        $officeMonitorConfig = GetConfig 'OfficeMonitor'
        if ($officeMonitorConfig -eq 'true')
        {
            $OfficeMonitor = get-process | where-object {$_.MainWindowTitle -eq 'OfficeMonitor'}
            if ($OfficeMonitor -eq $null)
            {
                write-host 'OfficeMonitor Not Running, Starting New Process...' -f Cyan
                $expression = "c:\windows\system32\cmd.exe /c start powershell -version 5 -Command { $PSScriptRoot\officemonitor.ps1 }"
                invoke-expression $expression
            }
        }
    }
}

function AddEvent($ownerId, $eventType)
{
	$eventDate = Get-Date
	$query = "INSERT INTO Event (OwnerId, EventDate, EventType) VALUES($ownerId, '$eventDate', '$eventType')"
	SqlQueryInsert $query
}

function PostToFlow($ownerId, $message)
{
    $notificationsValue = [String] (GetConfig('Notifications'))
    if ($notificationsValue -eq 'true')
    {
        $notificationsOn = $true
    }
    else
    {
        $notificationsOn = $false
    }
   
    $SqlQuery = "SELECT *
    FROM Source
	WHERE Source.Id = $ownerId 
	ORDER BY Source.Id ASC"
    $source = SqlQueryReturn($SqlQuery)
    
    $sourceDirectory = $source.ADHomeDirectory
    $sourceDirectory = $sourceDirectory.Replace('\','\\')
    $batchNumber = $source.BatchNumber
    $destinationSiteUrl = $source.OneDriveUrl
    $destinationLibrary = $source.DestinationLibrary
    $destinationFolder = $source.DestinationFolder
    $sourceId = $source.Id
    $email = $source.SamAccountName
    $url =  GetConfig('FlowHttpUrl')
    write-host $notificationsOn $url
    if ($notificationsOn -AND $url -ne $null)
    {
        $body = '{
	        "message":"' + $message + '"
	        ,"source_directory":"' + $sourceDirectory + '"
	        ,"batch_number":"' + $batchNumber + '"
	        ,"source_id":"' + $sourceId  + '"
	        ,"status":"success"
	        ,"destination_site_url":"' + $destinationSiteUrl + '"
	        ,"destination_library":"' + $destinationLibrary + '"
	        ,"destination_folder":"' + $destinationFolder + '"
	        ,"server":"' + $env:COMPUTERNAME + '"
	        ,"successes":"-1"
	        ,"errors":"-1"
	        ,"warnings":"-1"
	        ,"email":"' + $email + '"
        }'
        write-host $body
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -ContentType "application/json" -Method POST -Body $body
    }
}


function CreateSPMTErrorReport($ownerId)
{
    #Get SPMT Error Logs and Filter
    if ($global:SqlServer)
    {
        $query = 'SELECT TOP 1 * FROM MigrationJob WHERE OwnerId = ' + $ownerId + ''
    }
    else
    {
        $query = 'SELECT * FROM MigrationJob WHERE OwnerId = ' + $ownerId + ' LIMIT 1'
    }
    write-host $query -f Yellow
    $migrationJob = SqlQueryReturn($query)
    $global:Temp = $migrationJob
    $apiReportPath = $migrationJob.ApiReportPath
    if ($apiReportPath -ne $null -AND $apiReportPath -ne '')
    {
        $migrationJobId = $migrationJob.Id
        #Give SMPT Chance to finish writing log files
        #Start-Sleep -s 5
        $counter = 1
        $logFileFound = $true
        $logItems = @()
        $hashLog = @{}
        while($logFileFound)
        {
            $logFileName = ([String] $apiReportPath) + '\ItemFailureReport_R' + $counter.ToString() + '.csv'
            write-host 'Checking for ' $logFileName
            if (Test-Path $logFileName -PathType Leaf)
            {
                $logFileFound = $true
            }
            else
            {
                $logFileFound = $false
            }
            if ($logFileFound)
            {
                write-host 'Reading Errors In:' $logFileName
                $rows = Import-CSv $logFileName
                #$rows = Import-Excel $logFileName
                foreach($row in $rows)
                {
                    $path = $row.Source
                    $filename = $row.'Item name'
                    $message = $row.Message
                    $extension = $row.Extension
                    $server = $row.'Device name'
                    $category = $row.'Result category'
                    $status = $row.Status
                    $size = $row.'Item size (bytes)'
                    if ($size -ne $null)
                    {
                        $size = [Float] $size
                        $size = $size / 10000.0
                    }
                    $extension = $row.Extension
                    if ($category.Trim() -ne 'SCAN FILTER' -AND $status.Trim() -eq 'Failed')
                    {
                        $tempItem = New-Object -TypeName PsObject
	                    $tempItem | Add-Member -MemberType NoteProperty -Name 'Path' -Value $path
                        $tempItem | Add-Member -MemberType NoteProperty -Name 'Message' -Value $message
                        $tempItem | Add-Member -MemberType NoteProperty -Name 'Extension' -Value $extension
                        $tempItem | Add-Member -MemberType NoteProperty -Name 'Server' -Value $server
                        $tempItem | Add-Member -MemberType NoteProperty -Name 'Size' -Value $size
						$tempItem | Add-Member -MemberType NoteProperty -Name 'Type' -Value 'SPMT'
                        if ($hashLog[$path] -eq $null)
                        {
                            $hashLog[$path] = $tempItem
                        }
                    }
                }
            }
            $counter++
        }
		
		#add office errors to user report path
		$query = "SELECT * FROM MigrationFile WHERE Error IS NOT NULL AND Error <> '' AND Error <> ' ' AND OwnerId = $ownerId"
		$officeErrors = SqlQueryReturn -Query $query
		foreach($officeError in $officeErrors)
		{
		  $path = $officeError.Path
          $filename = $officeError.FileName
          $message = $officeError.Error
          $extension = $officeError.Extension
		  $size = $officeError.Size
		  
          $tempItem = New-Object -TypeName PsObject
		  $tempItem | Add-Member -MemberType NoteProperty -Name 'Path' -Value "$path\$filename"
          $tempItem | Add-Member -MemberType NoteProperty -Name 'Message' -Value $message
          $tempItem | Add-Member -MemberType NoteProperty -Name 'Extension' -Value $extension
          $tempItem | Add-Member -MemberType NoteProperty -Name 'Server' -Value $env:computername
          $tempItem | Add-Member -MemberType NoteProperty -Name 'Size' -Value $size
		  $tempItem | Add-Member -MemberType NoteProperty -Name 'Type' -Value 'Office'
		  $hashLog["$path\$filename"] = $tempItem
		}

        foreach ($key in $hashLog.Keys) 
        {
            $logItems = $logItems + $hashLog[$key]
        }
		
        $query = 'SELECT * FROM Source WHERE Id =' + $ownerId + ''
        #write-host $query -f Yellow
        $source = SqlQueryReturn -Query $query
        $email = $source.SamAccountName
        $email = $email.Split('@')[0]
    
        #update UserReportPath if there were errors
        if ($logItems.Length -gt 0)
        {
            $workingDirectory = Get-Location
            $timestamp = get-date -f _MM_dd_HH_mm_ss
            $sourceDirectory = $source.ADHomeDirectory
            $sourceDirectory = $sourceDirectory.Replace('\\','_')
            $sourceDirectory = $sourceDirectory.Replace('\','_')
            $sourceDirectory = $sourceDirectory.Replace('$','')
            $logFile = ([String] $workingDirectory) + '\logs\' + ([String] $sourceDirectory)  + ([String] $timestamp) + '.csv'
            $logFile = $logFile.Trim('_')
            $logItems | Export-Csv $logFile -NoTypeInformation
            $query = "UPDATE MigrationJob SET UserReportPath = '" +  $logFile + "' WHERE Id =" + $migrationJobId + ""
            write-host $query
            SqlQueryReturn -Query $query
            #give file chance to finish writing before it gets sent in an email
            Start-Sleep 5
        }
    }
}

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