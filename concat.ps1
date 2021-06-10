
$global:logItems = @()


$files = Get-ChildItem -path .\
foreach($file in $files)
{
    if ($file.FullName.ToLower().EndsWith(".csv") -eq $true)
    {
        $rows = Import-Csv -Path $file.FullName
        foreach($row in $rows)
        {
            $global:logItems = $global:logItems + $row
        }
    }
}

$logFile = ".\logs.csv"
$global:logItems  | Export-Csv $logFile -NoTypeInformation