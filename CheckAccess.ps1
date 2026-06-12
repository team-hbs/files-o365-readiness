param (
    [int]$batchNumber = -1,
	[string]$wave = '',
    [int]$sourceId = -1
)

. .\common.ps1


$sources = GetSources -batchnumber $batchNumber -wave $wave -ownerId $sourceId

$noAccessPaths = @()

foreach($source in $sources)
{
    $path = $source.ADHomeDirectory
    write-host . -NoNewline
    $tempPath = New-Object -TypeName PsObject
    $tempPath | Add-Member -MemberType NoteProperty -Name Path $path
    
    try {
        # Attempt to get directory contents
        $errs = @()
        #this one generates all the errors
        $overall = Get-ChildItem $path -recurse -ErrorVariable errs | Measure-Object -property length -sum
   
        $errorMessage = ''

        if ($errs.Length -gt 0)
        {
            #$err
            foreach($err in $errs)
            {
                $message = $err.Exception.Message
                if ($message.Contains('Could not find a part of') -eq $false)
                {
                    $tempPath = New-Object -TypeName PsObject
                    $tempPath | Add-Member -MemberType NoteProperty -Name Path $path
                    $tempPath | Add-Member -MemberType NoteProperty -Name Message $message
                    $noAccessPaths = $noAccessPaths + $tempPath
                }
            }
        }

	    $overallFileCount = $overall.Count
	
        write-host 'Item Count for' $path '=' $overallFileCount
        if ($items.Count -eq 0) {
            #Write-Output $path "Directory is empty."
        } else {
            #Write-Output $path "Directory is accessible and contains $($items.Count) item(s)."
        }
    }
    catch [System.Exception] 
    {
        $message = $_.Exception.Message
        if ($message.Contains('Could not find a part of') -eq $false)
        {
            Write-host $path "An error occurred: $($_.Exception.Message)"
            $tempPath = New-Object -TypeName PsObject
            $tempPath | Add-Member -MemberType NoteProperty -Name Path $path
            $tempPath | Add-Member -MemberType NoteProperty -Name Message $message
            $noAccessPaths = $noAccessPaths + $tempPath
        }
    }
}
$Global:noAccessPaths = $noAccessPaths
$logFile = '.\noaccesspaths_' + $batchNUmber  + '_.csv'
$noAccessPaths | Export-Csv -path $logFile -NoTypeInformation