Import-Module ActiveDirectory
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

$users = @()
$activeDirectoryUsers = Get-ADUser -Filter * -Properties HomeDirectory,Mail,Department | Where { $_.Enabled -eq $True}

foreach ($row in $activeDirectoryUsers)
{
    write-host '.' -NoNewline
    $tempUser = New-Object -TypeName PsObject
    $tempUser | Add-Member -MemberType NoteProperty -Name BatchNumber ''
    $tempUser | Add-Member -MemberType NoteProperty -Name Email $row.Mail
    $tempUser | Add-Member -MemberType NoteProperty -Name SourceDirectory $row.HomeDirectory
    $tempUser | Add-Member -MemberType NoteProperty -Name Department $row.Department
    $tempUser | Add-Member -MemberType NoteProperty -Name DestinationLibrary ''
    $tempUser | Add-Member -MemberType NoteProperty -Name DestinationFolder ''
    if ($row.Mail -ne '' -AND $row.Mail -ne $null)
    {
        $users = $users + $tempUser
    }
}

write-host ''

$users | Export-Csv ./ad_batch_template.csv -NoTypeInformation