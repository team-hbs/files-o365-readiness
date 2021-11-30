$hashOffices = @{}
$lifetimeLimit = 4
 
$Host.UI.RawUI.WindowTitle = 'OfficeMonitor'
 
while($true)
{
    $offices = get-process | Where-object {$_.ProcessName -like '*excel*' -OR $_.ProcessName -like '*word*' -OR $_.ProcessName -like '*Powerpnt*'}
    foreach($app in $offices)
    {
        if ($hashOffices[$app.Id] -eq $null)
        {
            $now = Get-Date
            $tempItem = New-Object -TypeName PsObject
            $tempItem | Add-Member -MemberType NoteProperty -Name 'RunningSince' -Value $now
            $tempItem | Add-Member -MemberType NoteProperty -Name 'Id' -Value $app.Id
            $tempItem | Add-Member -MemberType NoteProperty -Name 'Name' -Value $app.ProcessName
            $hashOffices[$app.Id] = $tempItem
            write-host ''
            write-host 'Logging Process' $app.ProcessName 'Id' $app.Id $now -f Cyan
        }
        else
        {
            $now = Get-Date
            $tempItem = $hashOffices[$app.Id]
            $runningSince = $tempItem.RunningSince
            $minutes = ($now - $runningSince).TotalMinutes
            if ($minutes -gt $lifetimeLimit)
            {
                write-host ''
                write-host 'Stopping Process' $tempItem.Name 'Id' $tempItem.Id  $now -f Yellow
                Stop-Process -Id $app.Id  -Confirm:$false -PassThru
                $hashOffices.Remove($app.Id)
            }
        }
    }
    #cleanup processes that closed on their own
    foreach($id in $hashOffices.Keys)
    {
        $tempItem = $hashOffices[$id]
        $runningSince = $tempItem.RunningSince
        $minutes = ($now - $runningSince).TotalMinutes
        if ($minutes -gt ($lifetimeLimit + 10))
        {
            $office = get-process | Where-object {$_.Id -like $id}
            if ($office -eq $null)
            {
                $hashOffices.Remove($tempItem.Id)
            }
        }
    }
    write-host '.' -NoNewline
    Start-Sleep -seconds 10
}