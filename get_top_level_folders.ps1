$path = '\\ts-rackstation\shared'
#$path = '\\ts-rackstation\previous_staff'

$folders = Get-ChildItem -Path $path -Directory
$folders | Export-Csv -path '.\shared.csv' -NoTypeInformation
#$folders | Export-Csv -path '.\previous_staff.csv' -NoTypeInformation

