param (
	[string] $path = '',
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


$sqlCounts = "SELECT count(Id) as FileCount, Sum(size) as FileSize FROM ScanFile WHERE ParentFolder like '{PARENT_FOLDER}%'"
	
$rows = Import-CSV -path $path

foreach($row in $rows)
{	
    $directory = $row.ADHomeDirectory
	$logFile = '.\file_counts.xlsx'
		
    $query = $sqlCounts.Replace('{PARENT_FOLDER}',$directory)
	$result = SqlQueryReturn($query)
  
	foreach($row in $result)
	{
                $tempItem = New-Object -TypeName PsObject
		        $tempItem | Add-Member -MemberType NoteProperty -Name 'Path' -Value $directory
                $tempItem | Add-Member -MemberType NoteProperty -Name 'FileCount' -Value $total.FileCount
                $tempItem | Add-Member -MemberType NoteProperty -Name 'Size' -Value $total.FileSize
                $fileCounts = $fileCounts + $tempitem
    }
 }
 $fileCounts | Export-Csv -path $logFile -NoTypeInformation
