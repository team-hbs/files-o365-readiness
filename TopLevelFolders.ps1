param (
	[string]$path = 'c:\projects'
)

$tempFolders = @()

$folders = Get-ChildItem -Directory -Path $path
foreach($folder in $folders)
{
    $tempFolder = New-Object -TypeName PsObject
    $tempFolder | Add-Member -MemberType NoteProperty -Name HomeDirectory  $folder.FullName
    $tempFolder | Add-Member -MemberType NoteProperty -Name Email  ''
    $tempFolders = $tempFolders + $tempFolder
}

$timestamp = get-date -f _MM_dd_HH_mm_ss
$logFile = ".\Batches" + $timestamp + ".csv"
$tempFolders | Export-Csv $logFile -NoTypeInformation