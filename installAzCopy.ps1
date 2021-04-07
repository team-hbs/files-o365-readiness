param (
    [Parameter()]
    [string]$InstallPath = "$PSScriptRoot"
)

# Cleanup Destination
<#
if (Test-Path $InstallPath) {
    Get-ChildItem $InstallPath | Remove-Item -Confirm:$false -Force
}
#>

# Zip Destination
$zip = "$InstallPath\AzCopy.Zip"

# Create the installation folder (eg. C:\AzCopy)
$null = New-Item -Type Directory -Path $InstallPath -Force

# Download AzCopy zip for Windows
Start-BitsTransfer -Source "https://aka.ms/downloadazcopy-v10-windows" -Destination $zip

# Expand the Zip file
Expand-Archive $zip $InstallPath -Force

# Move to $InstallPath
Get-ChildItem "$($InstallPath)\*\*" | Move-Item -Destination "$($InstallPath)\" -Force

#Cleanup - delete ZIP and old folder
Remove-Item $zip -Force -Confirm:$false
Get-ChildItem "$($InstallPath)\*" -Directory | ForEach-Object { Remove-Item $_.FullName -Recurse -Force -Confirm:$false }

# Add InstallPath to the System Path if it does not exist
<#
if ($env:PATH -notcontains $InstallPath) {
    $path = ($env:PATH -split ";")
    if (!($path -contains $InstallPath)) {
        $path += $InstallPath
        $env:PATH = ($path -join ";")
        $env:PATH = $env:PATH -replace ';;',';'
    }
    [Environment]::SetEnvironmentVariable("Path", ($env:path), [System.EnvironmentVariableTarget]::Machine)
}
#>
