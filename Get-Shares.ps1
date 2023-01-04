$servers = @(
 'madfsrv3'
,'lcfsrv'
,'bar-sv-fsrv'
,'oshfsrv'
,'chifsrv'
,'spfsrv'
,'lax-sv-fsrv'
,'erfsrv3'
)
<#
$servers = @(
 ,'erfsrv3'
)
#>
$global:sources = @()

foreach($server in $servers)
{
    try
    {
        write-host "Getting Shares for:" $server
        $remotesession = new-pssession -computername $server
        $shares = Invoke-Command -session $remotesession -ScriptBlock {  get-smbshare } 
        #exit-pssession

        foreach($share in $shares)
        {
            $path = $share.Path
            #$path
            if ($path.Contains(':\') -eq $true  -AND $share.Name.Contains('$') -eq $false)
            {
                $sharePath = "\\" + $server + "\" + $share.Name
                write-host $sharePath
                $source = New-Object -TypeName PsObject
                $source | Add-Member -MemberType NoteProperty -Name SourceDirectory $sharePath
                $source | Add-Member -MemberType NoteProperty -Name Server $server
                $global:sources = [Array] $global:sources + $source
            }
        }
    }
    catch
    {
            $line = $_.InvocationInfo.ScriptLineNumber.ToString()
            $message = $line + " " + $_.Exception.Message
            write-host $message -ForegroundColor Red
    }
}

$timestamp =  get-date -f _MM_dd_HH_mm_ss
$logFile = ".\fileshares_" + $timestamp + ".csv"


$global:sources | Export-Csv -path $logFile -NoTypeInformation