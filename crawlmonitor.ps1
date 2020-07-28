param (
    [string]$connectionString,
    [string]$email
)


Write-Host "Monitoring Scan" -ForegroundColor Red

$DataSource = $PSScriptRoot + "\FileToOneDrive.db"

$prevId = 0
$stuck = $false
while ($true) {
    Start-Sleep -Seconds 900
    $Id = 0
    if ($connectionString -ne "") {
        $query = "	SELECT TOP 1 *
                    FROM Files_OneDrive
                    ORDER BY Id Desc"
        $DataTable = $null
        $null = @(
            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	        $SqlConnection.ConnectionString = $connectionString
            
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd.CommandText = $query
			$SqlCmd.Connection = $SqlConnection
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCmd
			$DataSet = New-Object System.Data.DataSet
			$SqlAdapter.Fill($DataSet)
            $DataTable = $DataSet.Tables[0]
		)
        $Id = $DataTable.Id
        
	} else {
        $query = "	SELECT *
                    FROM Files_OneDrive
                    ORDER BY Id Desc
                    LIMIT 1"
        $result = Invoke-SqliteQuery -Query $query -DataSource $DataSource
        $Id = $result.Id
    }

    $smtp = "smtp.gmail.com"
    $from = "heartlandpowershellscripts@gmail.com"
    $username = "heartlandpowershellscripts@gmail.com"
    $password = ConvertTo-SecureString -String "heartland123" -AsPlainText -Force
    $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
    $to = $email.Split(";")

    if ($Id -eq $prevId) {
        if ($stuck -eq $false) {
            $stuck = $true
            Write-Host "Script Stuck!" -ForegroundColor Red
            
            $subject = "PowerShell Script Stuck!"
            $message = "PowerShell Script Stuck!"
            Send-MailMessage -To $to -From $from -Subject $subject -Body $message -SmtpServer $smtp -Credential $credential -UseSsl -Port 587 -DeliveryNotificationOption Never
        }
    } else {
        if ($stuck -eq $true) {
            $stuck = $false
            Write-Host "Script Resumed" -ForegroundColor Green

            $subject = "PowerShell Script Resumed"
            $message = "PowerShell Script Resumed"
            Send-MailMessage -To $to -From $from -Subject $subject -Body $message -SmtpServer $smtp -Credential $credential -UseSsl -Port 587 -DeliveryNotificationOption Never

        }
    }
    $prevId = $Id
}