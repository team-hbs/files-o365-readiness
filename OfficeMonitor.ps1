$hashOffices = @{}
$lifetimeLimit = 1
 
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
                write-host 'Stopping Process' $tempItem.Name 'Id' $app.Id '|' + $tempItem.Id "Minutes:" $minutes -f Yellow
                Stop-Process -Id $app.Id  -Confirm:$false -PassThru
                $hashOffices.Remove($app.Id)
                [gc]::collect()
                [gc]::WaitForPendingFinalizers()
            }
        }
    }
    #cleanup processes that closed on their own
    $removeIds = @()
    foreach($id in $hashOffices.Keys)
    {
        $tempItem = $hashOffices[$id]
        $runningSince = $tempItem.RunningSince
        #$minutes = ($now - $runningSince).TotalMinutes
        #if ($minutes -gt ($lifetimeLimit + 10))
        #{
            $office = get-process | Where-object {$_.Id -like $id}
            if ($office -eq $null)
            {
                #log ids because we can't modify collection while enumerating
                $removeIds = $removeIds + ([string] $tempItem.Id)
                #$hashOffices.Remove($tempItem.Id)
            }
        #}
    }
    foreach($id in $removeIds)
    {
        write-host 'Cleaning Up Id' $id -f Green
        $hashOffices.Remove([Int] $id)
    }
    write-host '.' -NoNewline
    Start-Sleep -seconds 10
}