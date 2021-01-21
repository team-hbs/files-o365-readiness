$rows = Import-csv .\user_homedrives.csv
foreach($row in $rows)
{
    if ($row.Directory -ne $null -AND $row.Directory.Trim() -ne '')
    {
        $homeDirectory = $row.Directory
        if ($homeDirectory.StartsWith('\\'))
        {
            #do nothing
        }
        elseif($homeDirectory.StartsWith('\'))
        {
            $homeDirectory = '\' + $homeDirectory
            $row.Directory = $homeDirectory
        }

    }
}

$newRows = @()
foreach($row in $rows)
{
    if ($row.Directory -ne $null -AND $row.Directory.Trim() -ne '')
    {
        $newRows += $row
    }
}
$rows = $newRows

$rows | Export-CSV .\user_homedrives_fixed.csv -NoTypeInformation