$global:unixEpoch = Get-Date -Date "01/01/1970"
$global:DataSource = $PSScriptRoot + "\FilesToO365.db"
$global:LastModifiedDate = $null

#function InitCrawl($ownerId, $email, $startPath, $doConvert)
function InitCrawl($ownerId, $startPath, $doConvert, $noOffice, $noLinks, $noSSN, $noCC, $lastModifiedDate)
{
	if ($lastModifiedDate -eq $null)
	{
		AddEvent -ownerId $ownerId -eventType 'ScanStarted'
	}
	else
	{
		AddEvent -ownerId $ownerId -eventType 'LastModifiedScanStarted'
	}
    if ($noOffice) {
        Write-Host "NO OFFICE MODE" -ForegroundColor Black -BackgroundColor White
    } else {
        $global:word = $null
        $global:excel = $null
        $global:powerpoint = $null
    }

	$filesUsersTableName = ""
	$filesTableName = ""
	if ($doConvert -eq $true)
	{
		$filesUsersTableName = "MigrationJob"
		$filesTableName = "MigrationFile"
	}
	else
	{
		$filesUsersTableName = "ScanJob"
		$filesTableName = "ScanFile"
	}
	
	$global:LastModifiedDate = $lastModifiedDate

    $StartPath = $path
    $ownerId = $ownerId
    $global:currentFileCount = 0
    $global:currentFileSize = 0
    $global:currentErrorCount = 0
    $StartPath = '' +  $StartPath + ''
    
    if ($StartPath.StartsWith('\\'))
    {
        $tempStartPath = 'FileSystem::' + $StartPath
    }
    else
    {
        $tempStartPath = $StartPath
    }
   
    write-host 'Getting overall count for ' $StartPath ' this may take a while....'
    
	#CHANGE THIS TO $true IF SCRIPT HANGS TOO LONG
	$skipInitialFileCount = $false
	if ($skipInitialFileCount)
	{
		$global:overallFileCount = 1000000000
	}
	else
	{
		$global:overallFileCount = (Get-ChildItem -LiteralPath $tempStartPath -Recurse -File | Measure-Object | Select-Object -Property Count).Count
	}
    
    CrawlFolder $StartPath $ownerId
	if ($lastModifiedDate -eq $null)
    {
        InsertUserEntry $ownerId $global:currentFileCount $global:currentFileSize $global:currentErrorCount
        UpdateExtensions $ownerId
        UpdateFileTotals $ownerId
        UpdateOfficeErrorTotals $ownerId
        UpdateMacroCount $ownerId
        UpdateLinkCount $ownerId
        UpdateOldOfficeCount $ownerId
        UpdatePathLengthCount  $ownerId
        UpdateNoAccessErrorTotals $ownerId
    	
        if ($doConvert -eq $true)
        {
            UpdateOfficeConversion $ownerId
        }
    }
    #clean up orphaned office instances
    if ($noOffice) {
        Write-Host "NO OFFICE MODE" -ForegroundColor Black -BackgroundColor White
    } else {
        #Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue
        #Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue
        #Stop-Process -Name "POWERPNT" -Force -ErrorAction SilentlyContinue
    }
	if ($lastModifiedDate -eq $null)
	{
		AddEvent -ownerId $ownerId -eventType 'ScanEnded'
	}
	else
	{
		AddEvent -ownerId $ownerId -eventType 'LastModifiedScanEnded'
	}
}



function UpdateMigrationStatus($ownerId, $status)
{
    try
    {
        $TempSqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $TempSqlCmd.Connection = $SqlConnection
        $TempSqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $query = "UPDATE $filesUsersTableName SET Migration = $status WHERE OwnerId = $ownerId"
        if ($global:SqlConnection.State -ne 1)
	    {
		    $global:SqlConnection.Open()
	    }
        $TempSqlCmd.CommandText = $query
	    write-host "update command:" $TempSqlCmd.CommandText
	    $rowsAffected = $TempSqlCmd.ExecuteNonQuery()
	    write-host "rows updated:" $rowsAffected
    }
	catch
	{
        $line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
		write-error $message
        $global:currentErrorCount++
        New-Object -TypeName PsObject -Property @{FileName="UpdateMigrationComplete";Message=$message;Path="";Query=""}
    }
}

function WaitForKeyPress($message)
{
    # Check if running Powershell ISE
    if ($psISE)
    {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("$message")
    }
    else
    {
        Write-Host "$message" -ForegroundColor Yellow
        $x = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

function CrawlFolder($path, $ownerId, $currentDepth)
{
	#check for stop file
	$workingDirectory = (Get-Location)
	$stopFile = $workingDirectory.Path + '\pause.txt'
	if (Test-Path -Path $stopFile -PathType Leaf)
	{
		WaitForKeyPress 'Delete pause.txt and Press Any Key To Continue'
	}

    try
    {
        if ($path.StartsWith('\\'))
        {
            $tempPath = 'FileSystem::' + $path
        }
        else
        {
            $tempPath = $path
        }
	    foreach ($file in Get-ChildItem -LiteralPath $tempPath -File -ErrorAction Continue)
	    {
		    $tempPath = $path
		    write-host $file.Name 
            #TODO: Check for modified vs ScanModified
          
            if ($global:LastModifiedDate -ne $null)
            {
                $lastModified = $file.LastWriteTime
				$created = $file.CreationTime
                if ($lastModified -gt (Get-Date $global:LastModifiedDate) -OR $created -gt (Get-Date $global:LastModifiedDate))
                {
		            InsertRow $file $path $ownerId $currentDepth
                }
            }
            else
            {
                InsertRow $file $path $ownerId $currentDepth
            }
	    }
        foreach ($folder in Get-ChildItem -LiteralPath $tempPath -Directory -ErrorAction Continue)
	    {
		    $subPath = $path + "\" + $folder.name
		    #Get-ChildItem $subPath -File
		    CrawlFolder $subPath $ownerId ($currentDepth + 1)
	    }
    }
    catch
    {
        if ($_.Exception.Message.Trim() -eq 'An unexpected network error occurred.')
        {
            InsertNoAccess $tempPath $ownerId
        }
    }
}



function CollectDocumentLinks2($ownerId)
{
	#query scanfile
	$query = "SELECT * FROM ScanFile WHERE OwnerId = $ownerId AND HasLink = 1"
	$result = SqlQueryReturn $query
    write-host 'Found' $result.Length 'Documents'
	foreach($row in $result)
	{
        $global:LinkCounter++
        $extension = $row.Extension
        $filePath = $row.Path + "\"  + $row.FileName
        $fileId = $row.Id
        $replaceExtension = $extension
        if ($extension -eq 'doc' -OR $extension -eq 'docx') 
        {
            try
            {
                #[gc]::collect()
                #[gc]::WaitForPendingFinalizers()
                $global:word = new-object -comobject word.application
                $global:word.Visible = $False
                $global:word.DisplayAlerts = [Enum]::Parse([Microsoft.Office.Interop.Word.WdAlertLevel],"wdAlertsNone")
                $global:word.AutomationSecurity = 'msoAutomationSecurityForceDisable'
                write-host 'Opening:' $filePath
                $opendoc = $global:word.documents.OpenNoRepairDialog($filePath,$false,$true,$false,'')
                write-host $global:LinkCounter 'Hyperlinks:' $opendoc.Hyperlinks.Count
                if ($opendoc.Hyperlinks.Count -gt 0)
                {
                    write-host 'FOUND WORD LINKS' -f WHITE
                    foreach($hyperlink in $opendoc.Hyperlinks)
                    {
                        $url = $hyperlink.Address
                        $created = Get-Date
                        if ($url -ne $null -AND $url.Trim() -ne '')
                        {
                            write-host 'LINK:' $url -f Cyan
                            $query = "INSERT INTO ScanLink (Url,OwnerId,FileId,Created) VALUES ('$url',$ownerId,$fileId,'$created')"
                            write-host $query
                            SqlQueryInsert -query $query
                        }
                    }
                }
            }
            catch
            {
               write-host $_.Exception.Message -f Yellow
            }
            if ($openDoc -ne $null)
            {
                $opendoc.close($false)
            }
            #Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue
        }  
        elseif ($extension -eq 'xls' -OR $extension -eq 'xlsx') 
        {
            try
            {
                #[gc]::collect()
                #[gc]::WaitForPendingFinalizers()
                $global:excel = new-object -comobject excel.application
                $global:excel.Visible = $False
                $global:excelSaveFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
                $global:excel.DisplayAlerts = $False;
                $global:excel.AutomationSecurity = 'msoAutomationSecurityForceDisable'
                $workBook  =  $global:excel.workbooks.open("$filePath", $false, $true, 5, "")
                foreach($worksheet in $workBook.worksheets)
                {
                    if ($worksheet.Hyperlinks.Count -gt 0)
                    {
                        write-host 'Found EXCEL links' -f white
                        foreach($hyperlink in $worksheet.Hyperlinks)
                        {
                            $url = $hyperlink.Address
                            $created = Get-Date
                            if ($url -ne $null -AND $url.Trim() -ne '')
                            {
                                write-host 'LINK:' $url -f Cyan
                                $query = "INSERT INTO ScanLink (Url,OwnerId,FileId,Created) VALUES ('$url',$ownerId,$fileId,'$created')"
                                write-host $query
                                SqlQueryInsert -query $query
                            }
                        }
                    }
                }
            }
            catch
            {
               write-host $_.Exception.Message -f Yellow
            }
            if ($workbook -ne $null)
            {
                $workBook.close($false);
            }
            #Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue
        }
        elseif ($extension -eq 'ppt' -OR $extension -eq 'pptx') 
        {
            try
            {
                #[gc]::collect()
                #[gc]::WaitForPendingFinalizers()
                $global:powerpoint = New-Object -ComObject PowerPoint.application
                $global:powerpointSaveFormat = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentation 
                $global:powerpoint.DisplayAlerts =  [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone
                $global:powerpoint.AutomationSecurity = 'msoAutomationSecurityForceDisable'

                $presentation = $global:powerpoint.Presentations.open("$filePath::password::", $true, $null, $false)
                foreach($slide in $presentation.Slides)
                {
                    if ($slide.Hyperlinks.Count -gt 0)
                    {
                        write-host 'Found powerpoint link' -f white
                        foreach($hyperlink in $slide.Hyperlinks)
                        {
                            $url = $hyperlink.Address
                            $created = Get-Date
                            if ($url -ne $null -AND $url.Trim() -ne '')
                            {
                                write-host 'LINK:' $url -f Cyan
                                $query = "INSERT INTO ScanLink (Url,OwnerId,FileId,Created) VALUES ('$url',$ownerId,$fileId,'$created')"
                                write-host $query
                                SqlQueryInsert -query $query
                            }
                        }
                    }
                }
                if ($presentation -ne $null)
                {
                    $presentation.close()
                }
            }
            catch
            {
               write-host $_.Exception.Message -f Yellow
            }
            #Stop-Process -Name "POWERPNT" -Force -ErrorAction SilentlyContinue
        }
    } 
}

function CollectDocumentLinks($ownerId)
{
	#query scanfile
	$query = "SELECT * FROM ScanFile WHERE OwnerId = $ownerId AND HasLink = 1"
	$result = SqlQueryReturn $query
   
	foreach($row in $result)
	{
        $extension = $row.Extension
        $filePath = $row.Path + "\"  + $row.FileName
        $replaceExtension = $extension
        if ($extension -eq 'doc' -OR $extension -eq 'docx') 
        {
            try 
            {
                #write-host 'Inspecting:' $filePath
                # Assemble paths
                $arrPath = $filePath.Split('\')
                $filename = $arrPath[$arrPath.Length - 1].ToLower().Replace(".$extension", '')
                $newFolder = "$PSScriptRoot\temp\$filename"
                $copyPath = "$newFolder\$filename.$extension"
                $zipPath = "$newFolder\$filename.zip"

                # Create temp folder
                if ((Test-Path -Path $newFolder) -eq $false) 
                {
                    New-Item -ItemType directory -Path $newFolder
                }

                # Make a copy of the file to work on
                Copy-Item $filePath -Destination $newFolder
    
                #Write-Host 'Renaming docx to zip'
                Rename-Item -Path $copyPath -NewName "$filename.zip" -Force
                Write-Host 'Unzipping Word document'
                Expand-Archive -LiteralPath $zipPath -DestinationPath $newFolder -Force

                #[xml]$xmlDoc = Get-Content "$newFolder\word\document.xml"
                [xml]$xmlDoc = Get-Content "$newFolder\word\_rels\document.xml.rels"
                #$nodes = $xmlDoc | Select-Xml -Xpath "//*[name()='w:hyperlink']"
                $nodes = $xmlDoc | Select-Xml -Xpath "//*[name()='Relationship']"
                $global:xmlDoc = $xmlDoc
                if($nodes -ne $null) 
                {
                    #Write-Host "REFERENCED LINKS FOUND"
				    foreach($node in $nodes)
                    {
                        if ($node.Node.TargetMode -eq 'External')
                        {
                            $url = $node.Node.Target
                            $created = Get-Date
                            write-host 'LINK:' $node.Node.Target -f Cyan
                            $query = "INSERT INTO ScanLink (Url,OwnerId,FileId,Created) VALUES ('$url',$ownerId,$fileId,'$created')"
                            SqlQueryInsert -query $query
                        }
                    }
                } 
                else
                {
                    #Write-Host "NO EMBEDDED LINKS FOUND" -f Yellow
                }
                Write-Host "Removing temp copy" $newFolder
                Remove-Item $newFolder -Recurse
            } 
            catch 
            {
                Write-Host "Error checking links" -f Yellow
            }
        }  
    } 
}
#function InsertUserEntry($email, $ownerId, $fileCount, $fileSize, $errorCount)
function InsertUserEntry($ownerId, $fileCount, $fileSize, $errorCount)
{
    $query = ''
    try
    {
		$today = Get-Date
        if ($global:SqlServer)
        {
            $created = (Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
            $query = "INSERT INTO  $filesUsersTableName  (OwnerId,FileCountDisk,FileSizeDisk,ErrorCount,CreatedDate) VALUES ($ownerId,$fileCount,$fileSize,$errorCount,'$created')"
            
        }
        else
        {
           $created = [int] (New-TimeSpan -Start $unixEpoch -End $today).TotalSeconds
           $query = "INSERT INTO  $filesUsersTableName  (OwnerId,FileCountDisk,FileSizeDisk,ErrorCount,CreatedDate) VALUES ($ownerId,$fileCount,$fileSize,$errorCount,$created)"
        }
        
        write-host "Query:" $query -ForegroundColor Green 	   
	    SqlQueryInsert($query)
    }
	catch
	{
	    $line = $_.InvocationInfo.ScriptLineNumber
		$message = $line.ToString() + " " + $_.Exception.Message
        New-Object -TypeName PsObject -Property @{FileName="InsertUserEntry";Message=$message;Query=$query}
        $global:currentErrorCount++
	}
}


function ConvertDocument($path, $file, $saveAs)
{
    #################
   
    #################
    $result = New-Object -TypeName psobject 
    $result | Add-Member -MemberType NoteProperty -Name HasMacro -Value $false
    $result | Add-Member -MemberType NoteProperty -Name Links -Value @()
    $result | Add-Member -MemberType NoteProperty -Name SSNMatches -Value @()
    $result | Add-Member -MemberType NoteProperty -Name CCMatches -Value @()
    $result | Add-Member -MemberType NoteProperty -Name ConvertMessage -Value ""
    $result | Add-Member -MemberType NoteProperty -Name ConvertSuccess -Value $false

    $filePath = $path + "\" + $file.Name
    $name = $file.Name
    write-host "Inspecting:" $filePath 
	$parts = $filePath.Split('.')
	$extension = $parts[$parts.length - 1]
	$baseFileName = $filePath.Replace("." + $extension, "")
	$converted = $false
	$message = ""
    $oldFormat = $false
	
	$global:word = $null
    $global:excel  = $null
    $global:powerpoint =  $null
	
	try
	{
       
		if ($extension -eq "doc" -OR $extension -eq "docx")
		{
            $oldFormat = $true
            if ($noOffice) {
                Write-Host "NO OFFICE MODE" -ForegroundColor Black -BackgroundColor White
            } else {
                if ($global:word -eq $null -OR $global:word.documents -eq $null)
                {
                    #[gc]::collect()
                    #[gc]::WaitForPendingFinalizers()
                    $global:word = new-object -comobject word.application
                    $global:word.Visible = $False
                    $global:word.DisplayAlerts = [Enum]::Parse([Microsoft.Office.Interop.Word.WdAlertLevel],"wdAlertsNone")
                    #new 7/27/19
                    #$global:excel.DisplayAlerts = $False;
                    $global:wordSaveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat],"wdFormatDocumentDefault")
                    $global:word.AutomationSecurity = 'msoAutomationSecurityForceDisable'
					#$global:word.AutoRecover.Enabled = $false
                }
                if ($extension -eq "docx")
                {
                    $opendoc = $null
                    if (!$noLinks -OR !$noSSN -OR !$noCC)
                    {
                        $opendoc = $global:word.documents.OpenNoRepairDialog($filePath,$false,$true,$false,'')
                    
                        if (!$noLinks)
                        {
                            if ($opendoc.Hyperlinks.Count -gt 0)
                            {
                                Write-Host 'FOUND WORD LINK' -f WHITE

                                foreach($hyperlink in $opendoc.Hyperlinks)
                                {
                                    $result.Links += $hyperlink.Address
                                }
                            }
                        }
                        if (!$noSSN)
                        {
                            foreach ($paragraph in $opendoc.Paragraphs) {
                                if ($paragraph.Range.Text -match "^(\D*|.*\D+)\d{3}[-|.\s]*\d{2}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                    Write-Host 'FOUND WORD SSN MATCH' -f WHITE
                                    $result.SSNMatches += ($Matches[0] -replace '\d', 'X')
                                }
                            }
                        }
                        if (!$noCC)
                        {
                            foreach ($paragraph in $opendoc.Paragraphs) {
                                if ($paragraph.Range.Text -match "^(\D*|.*\D+)\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                    Write-Host 'FOUND WORD CC MATCH' -f WHITE
                                    $result.CCMatches += ($Matches[0] -replace '\d', 'X')
                                }
                            }
                        }
                        $opendoc.close($false)
                    }
                }
                elseif ($extension -eq "doc")
                {
                    $testFilePath = $filePath + "x"
                    if ([System.IO.File]::Exists($testFilePath) -eq $false)
                    {
                        #$savename = $filePath.ToLower() -replace ".doc", ".docx"
                        $savename = $filePath.ToLower() + 'x'
                        #copy to local location
                        Write-Host "opening:" $filePath  
					    if ($doConvert)
					    {
						    write-host "Saving as :" $savename -ForegroundColor Cyan
					    }
                        try
                        {			    
                            #$opendoc = $global:word.documents.open($filePath,$false,$true)
                            #new 7/27/19
                            #$opendoc = $global:word.documents.OpenNoRepairDialog($filePath,$false,$true)
                            #new 2/5/21
                            $opendoc = $global:word.documents.OpenNoRepairDialog($filePath,$false,$true,$false,'')
                            if(!$noLinks)
                            {
                                write-host 'HYPERLINK COUNT:' 	$opendoc.Hyperlinks.Count			        
                                if ($opendoc.Hyperlinks.Count -gt 0)
                                {
                                    write-host 'FOUND WORD LINK' -f WHITE

                                    foreach($hyperlink in $opendoc.Hyperlinks)
                                    {
                                        $result.Links += $hyperlink.Address
                                    }
                                }
                            }
                            if (!$noSSN)
                            {
                                foreach ($paragraph in $opendoc.Paragraphs) {
                                    if ($paragraph.Range.Text -match "^(\D*|.*\D+)\d{3}[-|.\s]*\d{2}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                        Write-Host 'FOUND WORD SSN MATCH' -f WHITE
                                        $result.SSNMatches += ($Matches[0] -replace '\d', 'X')
                                    }
                                }
                            }
                            if (!$noCC)
                            {
                                foreach ($paragraph in $opendoc.Paragraphs) {
                                    if ($paragraph.Range.Text -match "^(\D*|.*\D+)\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                        Write-Host 'FOUND WORD CC MATCH' -f WHITE
                                        $result.CCMatches += ($Matches[0] -replace '\d', 'X')
                                    }
                                }
                            }
                            if ($saveAs -eq $true -AND $opendoc -ne $null)
                            {
                                $opendoc.saveas([ref]"$savename", [ref]$global:wordSaveFormat);
                                $converted = $true
                            }
						    $opendoc.close($false)
                            if ($opendoc -eq $null)
                            {
                                write-host "DOC IS NULL" -ForegroundColor yellow
                                throw "Sorry, we couldn't find your file. Was it moved, renamed, or deleted?"
                            }
                        }
                        catch
                        {
                            if ($_.Exception.Message.StartsWith("Sorry, we couldn't find your file. Was it moved, renamed, or deleted?"))
                            {
                                $tempFilePath = "c:\temp\" + $name
                        
                                $tempSaveName = $tempFilePath.ToLower() + 'x'
                                #copy to local location
                                Copy-Item $filePath -Destination $tempFilePath
                                #$opendoc = $global:word.documents.open($tempFilePath,$false,$true)
                                #new 7/27/19
                                #$opendoc = $global:word.documents.OpenNoRepairDialog($tempFilePath,$false,$true)
                                #new 2/5/21
                                $opendoc = $global:word.documents.OpenNoRepairDialog($tempFilePath,$false,$true,$false,'')
                                if(!$noLinks)
                                {
                                    write-host 'HYPERLINK COUNT:' 	$opendoc.Hyperlinks.Count	-f Cyan
                                    if ($opendoc.Hyperlinks.Count -gt 0)
                                    {
                                        write-host 'FOUND WORD LINK' -f WHITE

                                        foreach($hyperlink in $opendoc.Hyperlinks)
                                        {
                                            $result.Links += $hyperlink.Address
                                        }
                                    }
                                }
                                if (!$noSSN)
                                {
                                    foreach ($paragraph in $opendoc.Paragraphs) {
                                        if ($paragraph.Range.Text -match "^(\D*|.*\D+)\d{3}[-|.\s]*\d{2}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                            Write-Host 'FOUND WORD SSN MATCH' -f WHITE
                                            $result.SSNMatches += ($Matches[0] -replace '\d', 'X')
                                        }
                                    }
                                }
                                if (!$noCC)
                                {
                                    foreach ($paragraph in $opendoc.Paragraphs) {
                                        if ($paragraph.Range.Text -match "^(\D*|.*\D+)\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                            Write-Host 'FOUND WORD CC MATCH' -f WHITE
                                            $result.CCMatches += ($Matches[0] -replace '\d', 'X')
                                        }
                                    }
                                }
                                if ($saveAs)
                                {
                                    $opendoc.saveas([ref]"$tempSaveName", [ref]$global:wordSaveFormat);
                                }
                                $opendoc.close($false);
                                $newTempFilePath =  "c:\temp\" + $name  + "x" 
                                #copy back to original location
                                if ($saveAs -eq $true)
                                {
                                    Copy-Item ($newTempFilePath)  -Destination ($filePath + "x")
                                }
                                Remove-Item -Path $tempFilePath 
                                if ($saveAs -eq $true)
                                {
                                    Remove-Item -Path $newTempFilePath 
                                }
                                $converted = $true
                            }
                            else
                            {
                                throw $_
                            }
                        }
                        finally
                        {
                            #Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue
                        }
                    }
                    else
                    {
                        $oldFormat = $false
                    }
                }
            }
		}
		elseif ($extension -eq "xls" -OR $extension -eq 'xlsx')
		{
            $oldFormat = $true
            if ($noOffice) 
            {
                Write-Host "NO OFFICE MODE" -ForegroundColor Black -BackgroundColor White
            } 
            else 
            {
                if ($global:excel -eq $null -OR $global:excel.workbooks -eq $null)
                {
                    #[gc]::collect()
                    #[gc]::WaitForPendingFinalizers()
                    $global:excel = new-object -comobject excel.application
                    $global:excel.Visible = $False
                    $global:excelSaveFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
                    $global:excel.DisplayAlerts = $False;
                    $global:excel.AutomationSecurity = 'msoAutomationSecurityForceDisable'
					#$global:excel.AutoRecover.Enabled = $false
                }
                if ($extension -eq "xlsx")
                {
                    $workBook  =  $global:excel.workbooks.open("$filePath", $false, $true, 5, "")
                    if(!$noLinks)
                    {
                        foreach($worksheet in $workBook.worksheets)
                        {
                            if ($worksheet.Hyperlinks.Count -gt 0)
                            {
                                write-host 'Found EXCEL link' -f white

                                foreach($hyperlink in $worksheet.Hyperlinks)
                                {
                                    $result.Links += $hyperlink.Address
                                }
                            }    
                        }
                        
                    }
                    if (!$noSSN)
                    {
                        foreach ($worksheet in $workBook.Sheets) {
                            $rowCount = $worksheet.UsedRange[$worksheet.UsedRange.Count].Row
                            $columnCount = $worksheet.UsedRange[$worksheet.UsedRange.Count].Column
                        
                            for ($i = 1; $i -le $rowCount; $i++) {
                                for ($j = 1; $j -le $columnCount; $j++) {
                                    if ($worksheet.Cells.Item($i, $j).Text -match "^(\D*|.*\D+)\d{3}[-|.\s]*\d{2}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                        Write-Host 'FOUND EXCEL SSN MATCH' -f WHITE
                                        $result.SSNMatches += ($Matches[0] -replace '\d', 'X')
                                    }
                                }
                            }
                        }
                    }
                    if (!$noCC)
                    {
                        foreach ($worksheet in $workBook.Sheets) {
                            $rowCount = $worksheet.UsedRange[$worksheet.UsedRange.Count].Row
                            $columnCount = $worksheet.UsedRange[$worksheet.UsedRange.Count].Column

                            for ($i = 1; $i -le $rowCount; $i++) {
                                for ($j = 1; $j -le $columnCount; $j++) {
                                    if ($worksheet.Cells.Item($i, $j).Text -match "^(\D*|.*\D+)\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                        Write-Host 'FOUND EXCEL CC MATCH' -f WHITE
                                        $result.CCMatches += ($Matches[0] -replace '\d', 'X')
                                    }
                                }
                            }
                        }
                    }
                    $workBook.close($false);
                }
                elseif ($extension -eq "xls")
                {
                    $testFilePath = $filePath + "x"
                    if ([System.IO.File]::Exists($testFilePath) -eq $false)
                    {
                        try
                        {
                            $savename = $filePath.ToLower() + 'x'
                            $workBook  =  $global:excel.workbooks.open("$filePath", $false, $true, 5, "")
                            if(!$noLinks)
                            {
                                foreach($worksheet in $workBook.worksheets)
                                {
                                    if ($worksheet.Hyperlinks.Count -gt 0)
                                    {
                                        write-host 'Found EXCEL link' -f white

                                        foreach($hyperlink in $worksheet.Hyperlinks)
                                        {
                                            $result.Links += $hyperlink.Address
                                        }
                                    }
                                }
                            }
                            if (!$noSSN)
                            {
                                foreach ($worksheet in $workBook.Sheets) {
                                    $rowCount = $worksheet.UsedRange[$worksheet.UsedRange.Count].Row
                                    $columnCount = $worksheet.UsedRange[$worksheet.UsedRange.Count].Column
                                
                                    for ($i = 1; $i -le $rowCount; $i++) {
                                        for ($j = 1; $j -le $columnCount; $j++) {
                                            if ($worksheet.Cells.Item($i, $j).Text -match "^(\D*|.*\D+)\d{3}[-|.\s]*\d{2}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                                Write-Host 'FOUND EXCEL SSN MATCH' -f WHITE
                                                $result.SSNMatches += ($Matches[0] -replace '\d', 'X')
                                            }
                                        }
                                    }
                                }
                            }
                            if (!$noCC)
                            {
                                foreach ($worksheet in $workBook.Sheets) {
                                    $rowCount = $worksheet.UsedRange[$worksheet.UsedRange.Count].Row
                                    $columnCount = $worksheet.UsedRange[$worksheet.UsedRange.Count].Column

                                    for ($i = 1; $i -le $rowCount; $i++) {
                                        for ($j = 1; $j -le $columnCount; $j++) {
                                            if ($worksheet.Cells.Item($i, $j).Text -match "^(\D*|.*\D+)\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                                Write-Host 'FOUND EXCEL CC MATCH' -f WHITE
                                                $result.CCMatches += ($Matches[0] -replace '\d', 'X')
                                            }
                                        }
                                    }
                                }
                            }
                            if ($workbook.HasVBProject)
                            {
                                $result.HasMacro = $true
							    if ($saveAs -eq $true)
                                {
    								$savename = $savename -Replace ".xlsx", ".xlsm"
								    $workBook.saveas([ref]"$savename", [ref][Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled);
							    }
                            }
                            else
                            {
							    if ($saveAs -eq $true)
                                {
                                    $workBook.saveas([ref]"$savename", [ref]$global:excelSaveFormat);
                                }
                            }
                            $workBook.close($false);
                            $converted = $true
                        }
                        catch
                        {
                            if ($_.Exception.Message.StartsWith("Sorry, we couldn't find your file. Was it moved, renamed, or deleted?"))
                            {
                                $tempFilePath = "c:\temp\" + $name
                                #$tempSaveName  = ($tempFilePath).substring(0,($tempFilePath).lastindexOf("."))
                                $tempSaveName = $tempFilePath.ToLower() + 'x'
                                #copy to local location
                                Copy-Item $filePath -Destination $tempFilePath
                                $workBook  =  $global:excel.workbooks.open("$filePath", $false, $true, 5, "")
                                if(!$noLinks)
                                {
                                    foreach($worksheet in $workBook.worksheets)
                                    {
                                        if ($worksheet.Hyperlinks.Count -gt 0)
                                        {
                                            write-host 'Found EXCEL link' -f white

                                            foreach($hyperlink in $worksheet.Hyperlinks)
                                            {
                                                $result.Links += $hyperlink.Address
                                            }
                                        }
                                    }
                                }
                                if (!$noSSN)
                                {
                                    foreach ($worksheet in $workBook.Sheets) {
                                        $rowCount = $worksheet.UsedRange[$worksheet.UsedRange.Count].Row
                                        $columnCount = $worksheet.UsedRange[$worksheet.UsedRange.Count].Column
                                    
                                        for ($i = 1; $i -le $rowCount; $i++) {
                                            for ($j = 1; $j -le $columnCount; $j++) {
                                                if ($worksheet.Cells.Item($i, $j).Text -match "^(\D*|.*\D+)\d{3}[-|.\s]*\d{2}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                                    Write-Host 'FOUND EXCEL SSN MATCH' -f WHITE
                                                    $result.SSNMatches += ($Matches[0] -replace '\d', 'X')
                                                }
                                            }
                                        }
                                    }
                                }
                                if (!$noCC)
                                {
                                    foreach ($worksheet in $workBook.Sheets) {
                                        $rowCount = $worksheet.UsedRange[$worksheet.UsedRange.Count].Row
                                        $columnCount = $worksheet.UsedRange[$worksheet.UsedRange.Count].Column
                                        
                                        for ($i = 1; $i -le $rowCount; $i++) {
                                            for ($j = 1; $j -le $columnCount; $j++) {
                                                if ($worksheet.Cells.Item($i, $j).Text -match "^(\D*|.*\D+)\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                                    Write-Host 'FOUND EXCEL CC MATCH' -f WHITE
                                                    $result.CCMatches += ($Matches[0] -replace '\d', 'X')
                                                }
                                            }
                                        }
                                    }
                                }
                                if ($workbook.HasVBProject)
                                {
                                    $result.HasMacro = $true
								    if ($saveAs -eq $true)
								    {
									    $tempSaveName = $tempSaveName -Replace ".xlsx", ".xlsm"
									    $workBook.saveas([ref]"$tempSaveName", [ref][Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled);
								    }
                                }
                                else
                                {
                                    if ($saveAs -eq $true)
                                {
                                    $workBook.saveas([ref]"$tempSaveName", [ref]$global:excelSaveFormat);
                                }
                            }
                        
                            $workBook.close($false);
                            $converted = $true
                            #copy back to original location
                            if ($saveAs -eq $true)
                            {
                                Copy-Item ($newTempFilePath)  -Destination ($filePath + "x")
                            }
                            Remove-Item -Path $tempFilePath 
                            if ($saveAs -eq $true)
                            {
                                Remove-Item -Path $newTempFilePath 
                            }
                        }
                        else
                        {
                            throw $_
                        }
                    }
                    finally
                    {
                        #Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue
                    }
                }
                else
                {
                    $oldFormat = $false
                }
                }
            }
		}
		elseif ($extension -eq "ppt" -OR $extension -eq 'pptx')
		{
            $oldFormat = $true
            if ($noOffice)
            {
                Write-Host "NO OFFICE MODE" -ForegroundColor Black -BackgroundColor White
            }
            else
            {
                if ($global:powerpoint -eq $null -OR $global:powerpoint.Presentations -eq $null )
                {
                    #[gc]::collect()
                    #[gc]::WaitForPendingFinalizers()
                    $global:powerpoint = New-Object -ComObject PowerPoint.application
                    $global:powerpointSaveFormat = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentation 
                    $global:powerpoint.DisplayAlerts =  [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone
                    $global:powerpoint.AutomationSecurity = 'msoAutomationSecurityForceDisable'

                }
                if ($extension -eq 'pptx')
                {
                    $presentation = $global:powerpoint.Presentations.open("$filePath::password::", $true, $null, $false)
                    if(!$noLinks)
                    {
                        foreach($slide in $presentation.Slides)
                        {
                            if ($slide.Hyperlinks.Count -gt 0)
                            {
                                write-host 'Found powerpoint link' -f white

                                foreach($hyperlink in $slide.Hyperlinks)
                                {
                                    $result.Links += $hyperlink.Address
                                }
                            }
                        }
                    }
                    if (!$noSSN)
                    {
                        foreach ($slide in $presentation.Slides) {
                            foreach ($shape in $slide.Shapes) {
                                if ($shape.TextFrame.TextRange.Text -match "^(\D*|.*\D+)\d{3}[-|.\s]*\d{2}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                    Write-Host 'FOUND POWERPOINT SSN MATCH' -f WHITE
                                    $result.SSNMatches += ($Matches[0] -replace '\d', 'X')
                                }
                            }
                        }
                    }
                    if (!$noCC)
                    {
                        foreach ($slide in $presentation.Slides) {
                            foreach ($shape in $slide.Shapes) {
                                if ($shape.TextFrame.TextRange.Text -match "^(\D*|.*\D+)\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                    Write-Host 'FOUND POWERPOINT CC MATCH' -f WHITE
                                    $result.CCMatches += ($Matches[0] -replace '\d', 'X')
                                }
                            }
                        }
                    }
                    $presentation.close()
                }
                elseif ($extension -eq 'ppt')
                {
                    $testFilePath = $filePath + "x"
                    if ([System.IO.File]::Exists($testFilePath) -eq $false)
                    {
                        try
                        {
                            $presentation = $global:powerpoint.Presentations.open("$filePath::password::", $true, $null, $false)
                            #$savename = ($filePath).substring(0,($filePath).lastindexOf("."))
                            $savename = $filePath.ToLower() + 'x'
                            if(!$noLinks)
                            {
                                foreach($slide in $presentation.Slides)
                                {
                                    if ($slide.Hyperlinks.Count -gt 0)
                                    {
                                        write-host 'Found powerpoint link' -f white
                                    
                                        foreach($hyperlink in $slide.Hyperlinks)
                                        {
                                            $result.Links += $hyperlink.Address
                                        }
                                    }
                                }
                            }
                            if (!$noSSN)
                            {
                                foreach ($slide in $presentation.Slides) {
                                    foreach ($shape in $slide.Shapes) {
                                        if ($shape.TextFrame.TextRange.Text -match "^(\D*|.*\D+)\d{3}[-|.\s]*\d{2}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                            Write-Host 'FOUND POWERPOINT SSN MATCH' -f WHITE
                                            $result.SSNMatches += ($Matches[0] -replace '\d', 'X')
                                        }
                                    }
                                }
                            }
                            if (!$noCC)
                            {
                                foreach ($slide in $presentation.Slides) {
                                    foreach ($shape in $slide.Shapes) {
                                        if ($shape.TextFrame.TextRange.Text -match "^(\D*|.*\D+)\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                            Write-Host 'FOUND POWERPOINT CC MATCH' -f WHITE
                                            $result.CCMatches += ($Matches[0] -replace '\d', 'X')
                                        }
                                    }
                                }
                            }
                            if ($saveAs -eq $true)
                            {
                                $presentation.saveas([ref]"$savename", [ref]$global:powerpointSaveFormat);
                            }
                            $presentation.close();
                            $converted = $true
                        }
                        catch
                        {
                            if ($_.Exception.Message.StartsWith("Sorry, we couldn't find your file. Was it moved, renamed, or deleted?"))
                            {
                                $tempFilePath = "c:\temp\" + $name
                                #$tempSaveName  = ($tempFilePath).substring(0,($tempFilePath).lastindexOf("."))
                                $tempSaveName = $tempFilePath.ToLower() + 'x'
                                #copy to local location
                                Copy-Item $filePath -Destination $tempFilePath
                                $presentation = $global:powerpoint.Presentations.open("$tempFilePath::password::", $true, $null, $false)
                                if(!$noLinks)
                                {
                                    foreach($slide in $presentation.Slides)
                                    {
                                        if ($slide.Hyperlinks.Count -gt 0)
                                        {
                                            write-host 'Found powerpoint link' -f white
                                        
                                            foreach($hyperlink in $slide.Hyperlinks)
                                            {
                                                $result.Links += $hyperlink.Address
                                            }
                                        }
                                    }
                                }
                                if (!$noSSN)
                                {
                                    foreach ($slide in $presentation.Slides) {
                                        foreach ($shape in $slide.Shapes) {
                                            if ($shape.TextFrame.TextRange.Text -match "^(\D*|.*\D+)\d{3}[-|.\s]*\d{2}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                                Write-Host 'FOUND POWERPOINT SSN MATCH' -f WHITE
                                                $result.SSNMatches += ($Matches[0] -replace '\d', 'X')
                                            }
                                        }
                                    }
                                }
                                if (!$noCC)
                                {
                                    foreach ($slide in $presentation.Slides) {
                                        foreach ($shape in $slide.Shapes) {
                                            if ($shape.TextFrame.TextRange.Text -match "^(\D*|.*\D+)\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}[-|.\s]*\d{4}(\D*|\D+.*)$") {
                                                Write-Host 'FOUND POWERPOINT CC MATCH' -f WHITE
                                                $result.CCMatches += ($Matches[0] -replace '\d', 'X')
                                            }
                                        }
                                    }
                                }
                                if ($saveAs -eq $true)
                                {
                                    $presentation.saveas([ref]"$tempSaveName", [ref]$global:powerpointSaveFormat);
                                }
                                $presentation.close();
                                if ($saveAs -eq $true)
                                {
                                    Copy-Item ($newTempFilePath)  -Destination ($filePath + "x")
                                }
                                Remove-Item -Path $tempFilePath 
                                if ($saveAs)
                                {
                                    Remove-Item -Path $newTempFilePath 
                                }
                                $converted = $true
                            }
                            else
                            {
                                throw $_
                            }
                        }
                        finally
                        {
                            #Stop-Process -Name "POWERPNT" -Force -ErrorAction SilentlyContinue
                        }
                    }
                    else
                    {
                        $oldFormat = $false
                    }
                }
            }
        }
	}
	catch
	{
		$line = $_.InvocationInfo.ScriptLineNumber
		$message = $line.ToString() + ":" + $_.Exception.Message
       
        write-host $message  -ForegroundColor Red 
        $convertMessage = $message
	}

	if ($converted)
	{
		#$newFileName = $baseFileName + "_old." + $extension
		#Rename-Item -Path $filePath -NewName $newFileName
		New-Object -TypeName PsObject -Property @{Path=$filePath;Message="Document Converted Successfully";Converted=$converted;}
	}
	else
	{
        if ($oldFormat -eq $true)
        {
		    New-Object -TypeName PsObject -Property @{Path=$filePath;Message=$message;Converted=$converted;}
        }
	}
    $result.ConvertMessage = $convertMessage 
    $result.ConvertSuccess = $converted

    <#
    if ($extension -eq 'docx') {
        try {
            # Assemble paths
            $arrPath = $filePath.Split('\')
            $filename = $arrPath[$arrPath.Length - 1].ToLower().Replace('.docx', '')
            $newFolder = "$PSScriptRoot\temp\$filename"
            $copyPath = "$newFolder\$filename.docx"
            $zipPath = "$newFolder\$filename.zip"

            # Create temp folder
            if ((Test-Path -Path $newFolder) -eq $false) {
                New-Item -ItemType directory -Path $newFolder
            }

            # Make a copy of the file to work on
            Copy-Item $filePath -Destination $newFolder

            Write-Host 'Renaming docx to zip'
            Rename-Item -Path $copyPath -NewName "$filename.zip" -Force
            Write-Host 'Unzipping Word document'
            Expand-Archive -LiteralPath $zipPath -DestinationPath $newFolder -Force

            [xml]$xmlDoc = Get-Content "$newFolder\word\document.xml"
            $nodes = $xmlDoc | Select-Xml -Xpath "//*[name()='w:hyperlink']"

            if($nodes -ne $null) {
                Write-Host "EMBEDDED LINKS FOUND"
                $result.HasLink = $true
            } else {
                Write-Host "NO EMBEDDED LINKS FOUND"
            }
        } catch {
            Write-Host "Error checking links"
        }
        Write-Host "Removing temp copy"
        Remove-Item $newFolder -Recurse
    
    } elseif ($extension -eq 'xlsx') {
        try {
            # Assemble paths
            $arrPath = $filePath.Split('\')
            $filename = $arrPath[$arrPath.Length - 1].ToLower().Replace('.xlsx', '')
            $newFolder = "$PSScriptRoot\temp\$filename"
            $copyPath = "$newFolder\$filename.xlsx"
            $zipPath = "$newFolder\$filename.zip"

            # Create temp folder
            if ((Test-Path -Path $newFolder) -eq $false) {
                New-Item -ItemType directory -Path $newFolder
            }

            # Make a copy of the file to work on
            Copy-Item $filePath -Destination $newFolder

            Write-Host 'Renaming xlsx to zip'
            Rename-Item -Path $copyPath -NewName "$filename.zip" -Force
            Write-Host 'Unzipping Excel document'
            Expand-Archive -LiteralPath $zipPath -DestinationPath $newFolder -Force

            $worksheets = Get-ChildItem -Path "$newFolder\xl\worksheets" -File


            $embeddedFound = $false
            foreach ($worksheet in $worksheets) {
                [xml]$xmlDoc = Get-Content $worksheet.FullName
                $nodes = $xmlDoc | Select-Xml -Xpath "//*[name()='hyperlink']"
                if($nodes -ne $null) {
                    $embeddedFound = $true
                } 
            }

            if($embeddedFound) {
                Write-Host "EMBEDDED LINKS FOUND"
                $result.HasLink = $true
            } else {
                Write-Host "NO EMBEDDED LINKS FOUND"
            }
        } catch {
            Write-Host "Error checking links"
        }
        Write-Host "Removing temp copy"
        Remove-Item $newFolder -Recurse

    } elseif ($extension -eq 'pptx') {
        try {
            # Assemble paths
            $arrPath = $filePath.Split('\')
            $filename = $arrPath[$arrPath.Length - 1].ToLower().Replace('.pptx', '')
            $newFolder = "$PSScriptRoot\temp\$filename"
            $copyPath = "$newFolder\$filename.pptx"
            $zipPath = "$newFolder\$filename.zip"

            # Create temp folder
            if ((Test-Path -Path $newFolder) -eq $false) {
                New-Item -ItemType directory -Path $newFolder
            }

            # Make a copy of the file to work on
            Copy-Item $filePath -Destination $newFolder

            Write-Host 'Renaming pptx to zip'
            Rename-Item -Path $copyPath -NewName "$filename.zip" -Force
            Write-Host 'Unzipping PowerPoint document'
            Expand-Archive -LiteralPath $zipPath -DestinationPath $newFolder -Force

            $slides = Get-ChildItem -Path "$newFolder\ppt\slides" -File


            $embeddedFound = $false
            foreach ($slide in $slides) {
                [xml]$xmlDoc = Get-Content $slide.FullName
                $nodes = $xmlDoc | Select-Xml -Xpath "//*[name()='a:hlinkClick']"
                if($nodes -ne $null) {
                    $embeddedFound = $true
                } 
            }

            if($embeddedFound) {
                Write-Host "EMBEDDED LINKS FOUND"
                $result.HasLink = $true
            } else {
                Write-Host "NO EMBEDDED LINKS FOUND"
            }
        } catch {
            Write-Host "Error checking links"
        }
        Write-Host "Removing temp copy"
        Remove-Item $newFolder -Recurse

    }
    #>
    return $result
}

function UpdateExtensions($ownerId)
{
    try
    {
        $query = "select Extension, count(*) as count from $filesTableName Where OwnerId = $ownerId AND Extension IS NOT NULL GROUP BY Extension"
		$Result = SqlQueryReturn($query)
 
		Foreach ($row in $Result) 
        {
            if ($extensions -ne "")
            {
                $extensions += ","
            }
            $extension = $row.Extension
            $extensionCount = $row.count
            $extensions += $extension + ":" + $extensionCount
        }
        $query = "UPDATE  $filesUsersTableName SET Extensions = '$extensions' WHERE OwnerId = $ownerId"
        $rowsAffected = SqlQueryInsert($query)
    }
	catch
	{
    	$line = $_.InvocationInfo.ScriptLineNumber
		$message = ([string] $line) + " " + $_.Exception.Message
		write-error $message
        $global:currentErrorCount++
        New-Object -TypeName PsObject -Property @{FileName="UpdateExtensions";Message=$message;Path="";Query=""}
	}
}
function UpdateNoAccessErrorTotals($ownerId)
{
    try
    {
        $query = "select count(*) as count from  $filesTableName  where Error = 'No Access' AND OwnerId = $ownerId"
        $Result = SqlQueryReturn($query)
        $officeErrorCount = 0
        Foreach ($row in $Result) 
        {
            $officeErrorCount = $row.count
        }
        $query = "UPDATE  $filesUsersTableName  SET NoAccessCount = $officeErrorCount WHERE OwnerId = $ownerId"
	    $rowsAffected = SqlQueryInsert($query)
	    }
	catch
	{
        $line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
		write-error $message
        $global:currentErrorCount++
        New-Object -TypeName PsObject -Property @{FileName="UpdateNoAccessErrorTotals";Message=$message;Path="";Query=""}

	}
}

function UpdateOfficeErrorTotals($ownerId)
{
   $officeErrorCount = 0
  try
    {
        $query = "select count(*) as count from $filesTableName where Error IS NOT NULL AND Error <> '' AND Error <> ' ' AND Error <> 'No Access' AND OwnerId = $ownerId"
        $Result = SqlQueryReturn($query)
     
        Foreach ($row in $Result) 
        {
            $officeErrorCount = $row.count
        }
        $query = "UPDATE  $filesUsersTableName  SET OfficeErrorCount = $officeErrorCount WHERE OwnerId = $ownerId"
        $rowsAffected = SqlQueryInsert($query)
    }
	catch
	{
        $line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
		write-error $message
        $global:currentErrorCount++
        New-Object -TypeName PsObject -Property @{FileName="UpdateOfficeErrorTotals";Message=$message;Path="";Query=""}

	}
    

}

function UpdateMacroCount($ownerId)
{
  try
    {
        $query = "select count(*) as count from $filesTableName where HasMacro = 1 AND OwnerId = $ownerId"
		$Result = SqlQueryReturn($query)
        $macroCount = 0
        Foreach ($row in $Result) 
        {
            $macroCount = $row.count
        }
        $query = "UPDATE  $filesUsersTableName  SET MacroCount = $macroCount WHERE OwnerId = $ownerId"
        $rowsAffected = SqlQueryInsert($query)
	 }
	catch
	{
        $line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
		write-error $message
        $global:currentErrorCount++
        New-Object -TypeName PsObject -Property @{FileName="UpdateMacroCount";Message=$message;Path="";Query=""}

	}

}

function UpdateLinkCount($ownerId)
{
  try
    {
        $query = "select count(*) as count from $filesTableName where HasLink = 1 AND OwnerId = $ownerId"
		$Result = SqlQueryReturn($query)
        $linkCount = 0
        Foreach ($row in $Result) 
        {
            $linkCount = $row.count
        }
        $query = "UPDATE  $filesUsersTableName  SET LinkCount = $linkCount WHERE OwnerId = $ownerId"
        $rowsAffected = SqlQueryInsert($query)
	 }
	catch
	{
        $line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
		write-error $message
        $global:currentErrorCount++
        New-Object -TypeName PsObject -Property @{FileName="UpdateLinkCount";Message=$message;Path="";Query=""}

	}

}

function UpdateOldOfficeCount($ownerId)
{
  try
    {
        $query = "select count(*) as count from $filesTableName where (Extension = 'doc' OR Extension = 'xls' OR Extension = 'ppt') AND OwnerId = $ownerId"
       	$Result = SqlQueryReturn($query)
        $oldOfficeCount = 0
        Foreach ($row in $Result) 
        {
            $oldOfficeCount = $row.count
        }
        $query = "UPDATE  $filesUsersTableName  SET OldOfficeCount = $oldOfficeCount WHERE OwnerId = $ownerId"
        $rowsAffected = SqlQueryInsert($query)
    }
	catch
	{
        $line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
		write-error $message
        $global:currentErrorCount++
        New-Object -TypeName PsObject -Property @{FileName="UpdatOldOfficeCount";Message=$message;Path="";Query=""}

	}

}

function UpdatePathLengthCount($ownerId)
{
  try
    {
        $query = "select count(*) as count from $filesTableName where PathLength >= 218 AND OwnerId = $ownerId"
        $Result = SqlQueryReturn($query)
        $pathLengthCount = 0
        Foreach ($row in $Result) 
        {
            $pathLengthCount = $row.count
        }
        $query = "UPDATE  $filesUsersTableName  SET PathLengthCount = $pathLengthCount WHERE OwnerId = $ownerId"
        $rowsAffected = SqlQueryInsert($query)
    }
	catch
	{
        $line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
		write-error $message
        $global:currentErrorCount++
        New-Object -TypeName PsObject -Property @{FileName="UpdatePathLengthCount";Message=$message;Path="";Query=""}

	}

}

function UpdateFileTotals($ownerId)
{
    try
    {
        $query = "select count(*) as count, Sum(CAST(size as float)) as size from $filesTableName Where OwnerId = $ownerId"
        $Result = SqlQueryReturn($query)
        $extensions = ""
        $fileCountCrawl = 0
        $fileSizeCrawl = 0
        Foreach ($row in $Result) 
        {
            $fileCountCrawl = $row.count
            $fileSizeCrawl = $row.size
        }

        if ($fileSizeCrawl -eq $null -OR  ([DBNull]::Value).Equals($fileSizeCrawl))
        {
            $fileSizeCrawl = 0
        }
        $query = "UPDATE $filesUsersTableName  SET FileCountCrawl = $fileCountCrawl, FileSizeCrawl = $fileSizeCrawl WHERE OwnerId = $ownerId"
        $rowsAffected = SqlQueryInsert($query)
    }
	catch
	{
        $line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
		write-error $message
        $global:currentErrorCount++
        New-Object -TypeName PsObject -Property @{FileName="UpdateFileTotals";Message=$message;Path="";Query=""}
    }
}

function UpdateOfficeConversion($ownerId)
{
    try
    {
        $query = "UPDATE $filesUsersTableName  SET OfficeConversion = 1 WHERE OwnerId = $ownerId"
        $Result = SqlQueryInsert($query)
    }
	catch
	{
        $line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
		write-error $message
        $global:currentErrorCount++
        New-Object -TypeName PsObject -Property @{FileName="UpdateOfficeConversion";Message=$message;Path="";Query=""}
    }
}

function InsertNoAccess($path, $ownerId)
{
        $query = "INSERT INTO $filesTableName (Location, OwnerId, Error) VALUES ('$path', '$OwnerId','No Access')"
		$rowsAffected = SqlQueryInsert($query)
}

function InsertRow($file, $path, $ownerId, $currentDepth)
{

    $convertMessage = ""
	
    $saveAs = ($doConvert -eq $true)
    $result = $null
    try
    {
         $result = ConvertDocument $path $file $saveAs
    }
    catch
    {
        $result = New-Object -TypeName psobject 
        $result | Add-Member -MemberType NoteProperty -Name HasMacro -Value $false
        $result | Add-Member -MemberType NoteProperty -Name Links -Value @()
        $result | Add-Member -MemberType NoteProperty -Name SSNMatches -Value @()
        $result | Add-Member -MemberType NoteProperty -Name CCMatches -Value @()
        $result | Add-Member -MemberType NoteProperty -Name ConvertMessage -Value $line.ToString() + ":" + $_.Exception.Message
        $result | Add-Member -MemberType NoteProperty -Name ConvertSuccess -Value $false

    }
    $path = $path -Replace "'", "''"
    $tempFilePath = $path + "\" + $file.Name
    $tempFilePathLength = $tempFilePath.Length

    $convertSuccess = $result.ConvertSuccess
    $convertMessage = $result.ConvertMessage
    $hasMacroValue = 0
    $hasLinkValue = 0
    $hasSSNValue = 0
    $hasCCValue = 0
    $convertSuccessValue = 0

    if ($result.HasMacro -eq $true)
    {
        $hasMacroValue = 1
    }
    if ($result.Links.Length -gt 0)
    {
        $hasLinkValue = 1
    }
    if ($result.SSNMatches.Length -gt 0)
    {
        $hasSSNValue = 1
    }
    if ($result.CCMatches.Length -gt 0)
    {
        $hasCCValue = 1
    }
    if ($result.ConvertSuccess -eq $true)
    {
        $convertSuccessValue = 1
    }

	$tempFilePath = $path + "\" + $file.Name
    $tempFilePathLength = $tempFilePath.Length
	
    $global:currentFileSize +=  ($file.length / 1000000.0)
    $global:currentFileCount += 1
    Write-Progress -Id 1 -Activity "Crawl: $startPath" -Status "Progress: $global:currentFileCount / $global:overallFileCount Files" -PercentComplete ($global:currentFileCount / $global:overallFileCount * 100)
	$query = ''
	try{
	
	

        $tempParentFolderCurrent = $path.ToLower().Trim()

	    $replacePaths = @(
        "\\"
        "\\",
        "\\"
        )

        foreach ($term in $replacePaths)
        {
            $tempParentFolderCurrent =  $tempParentFolderCurrent.Replace($term, "\")
        }
       
     	$arrPathSplits = $tempParentFolderCurrent.Split('\')
	
		$FileName = $file.Name -Replace "\'","''"
		#HERE STUFF
	    $Created = $file.CreationTime
		$createdSeconds = (New-TimeSpan -Start $unixEpoch -End $Created).TotalSeconds
		$Modified = $file.LastWriteTime
        $modifiedSeconds = (New-TimeSpan -Start $unixEpoch -End $Modified).TotalSeconds
		try
        {
		    $Author = $file.GetAccessControl().Owner
        }
        catch
        {
            $author = $null
        }
        $Extension = $file.Name.Split('.')[$file.Name.Split('.').length - 1]
        $Size = [decimal]($file.length / 1000000.0)

		$OwnerId = $ownerId
		$Ignore = $false
		$Path = $path
		$FolderDepth = $currentDepth
		$Location = $Path + "\" + $FileName 
		
		if ($arrPathSplits.length -gt 1) { $Folder01 = $arrPathSplits[1].Trim().ToLower()}
		if ($arrPathSplits.length -gt 2) { $Folder02 = $arrPathSplits[2].Trim().ToLower()}
		if ($arrPathSplits.length -gt 3) { $Folder03 = $arrPathSplits[3].Trim().ToLower()}
		if ($arrPathSplits.length -gt 4) { $Folder04 = $arrPathSplits[4].Trim().ToLower()}
		if ($arrPathSplits.length -gt 5) { $Folder05 = $arrPathSplits[5].Trim().ToLower()}
		if ($arrPathSplits.length -gt 6) { $Folder06 = $arrPathSplits[6].Trim().ToLower()}
		if ($arrPathSplits.length -gt 7) { $Folder07 = $arrPathSplits[7].Trim().ToLower()}
		if ($arrPathSplits.length -gt 8) { $Folder08 = $arrPathSplits[8].Trim().ToLower()}
		if ($arrPathSplits.length -gt 9) { $Folder09 = $arrPathSplits[9].Trim().ToLower()}
		if ($arrPathSplits.length -gt 10) { $Folder10 = $arrPathSplits[10].Trim().ToLower()}
		if ($arrPathSplits.length -gt 11) { $Folder11 = $arrPathSplits[11].Trim().ToLower()}
		if ($arrPathSplits.length -gt 12) { $Folder12 = $arrPathSplits[12].Trim().ToLower()}
		if ($arrPathSplits.length -gt 13) { $Folder13 = $arrPathSplits[13].Trim().ToLower()}
		if ($arrPathSplits.length -gt 14) { $Folder14 = $arrPathSplits[14].Trim().ToLower()}
		if ($arrPathSplits.length -gt 15) { $Folder15 = $arrPathSplits[15].Trim().ToLower()}
		if ($arrPathSplits.length -gt 16) { $Folder16 = $arrPathSplits[16].Trim().ToLower()}
		if ($arrPathSplits.length -gt 17) { $Folder17 = $arrPathSplits[17].Trim().ToLower()}
		if ($arrPathSplits.length -gt 18) { $Folder18 = $arrPathSplits[18].Trim().ToLower()}
		if ($arrPathSplits.length -gt 19) { $Folder19 = $arrPathSplits[19].Trim().ToLower()}
		if ($arrPathSplits.length -gt 20) { $Folder20 = $arrPathSplits[20].Trim().ToLower()}
		
   
        $scanCreatedDate = (Get-Date).ToString('yyyy/MM/dd HH:mm:ss')

        $RelativeFolder = $tempParentFolderCurrent.Replace($StartPath.ToLower(), "")
		if ($global:SqlServer)
        {
            $createdValue = (Get-Date $Created).ToString('yyyy-MM-dd HH:mm:ss')
            $modifiedValue = (Get-Date $Modified).ToString('yyyy-MM-dd HH:mm:ss')
            $query = "INSERT INTO $filesTableName (FileName, Created, Modified, Author, Extension, Size, OwnerId, Ignore, Path, FolderDepth, ParentFolder, RelativeFolder,OfficeOpen,Error,PathLength, HasMacro, HasLink, HasSSN, HasCC,"
		    $query += "Folder01, Folder02, Folder03, Folder04, Folder05, Folder06, Folder07, Folder08, Folder09, Folder10, Folder11, Folder12, Folder13, Folder14, Folder15, Folder16, Folder17, Folder18, Folder19, Folder20,ScanCreatedDate) "
		    $query += " VALUES ('$FileName', '$createdValue', '$modifiedValue', '$Author', '$Extension', '$Size', $ownerId, '$Ignore', '$Path', '$FolderDepth','$tempParentFolderCurrent', '$RelativeFolder',$convertSuccessValue,'$convertMessage',$tempFilePathLength,$hasMacroValue,$hasLinkValue,$hasSSNValue,$hasCCValue,"
		    $query += " '$Folder01', '$Folder02', '$Folder03', '$Folder04', '$Folder05', '$Folder06', '$Folder07', '$Folder08', '$Folder09', '$Folder10', '$Folder11', '$Folder12', '$Folder13', '$Folder14', '$Folder15', '$Folder16', '$Folder17', '$Folder18', '$Folder19', '$Folder20','$scanCreatedDate')"
        }
        else
        {
            $query = "INSERT INTO $filesTableName (FileName, Created, Modified, Author, Extension, Size, OwnerId, Ignore, Path, FolderDepth, ParentFolder, RelativeFolder,OfficeOpen,Error,PathLength, HasMacro, HasLink, HasSSN, HasCC,"
		    $query += "Folder01, Folder02, Folder03, Folder04, Folder05, Folder06, Folder07, Folder08, Folder09, Folder10, Folder11, Folder12, Folder13, Folder14, Folder15, Folder16, Folder17, Folder18, Folder19, Folder20,ScanCreatedDate) "
		    $query += " VALUES ('$FileName', '$createdSeconds', '$modifiedSeconds', '$Author', '$Extension', '$Size', $ownerId, '$Ignore', '$Path', '$FolderDepth','$tempParentFolderCurrent', '$RelativeFolder',$convertSuccessValue,'$convertMessage',$tempFilePathLength,$hasMacroValue,$hasLinkValue,$hasSSNValue,$hasCCValue,"
		    $query += " '$Folder01', '$Folder02', '$Folder03', '$Folder04', '$Folder05', '$Folder06', '$Folder07', '$Folder08', '$Folder09', '$Folder10', '$Folder11', '$Folder12', '$Folder13', '$Folder14', '$Folder15', '$Folder16', '$Folder17', '$Folder18', '$Folder19', '$Folder20','$scanCreatedDate')"
        }
		
		

        write-host $query -ForegroundColor Green
		$rowsAffected = SqlQueryInsert($query)

        if (!$noLinks)
        {
            if ($hasLinkValue)
            {
                $query = "SELECT Id FROM $filesTableName WHERE FileName = '$FileName' AND Path = '$Path'"
                $sqlResult = SqlQueryReturn($query)
                foreach ($row in $sqlResult)
                {
                    $fileId = $row.Id
                }

                foreach ($link in $result.Links)
                {
                    if ($link -ne $null -AND $link.Trim() -ne '')
                    {
                        $created = Get-Date
                        $query = "INSERT INTO ScanLink (Url,OwnerId,FileId,Created) VALUES ('$link',$ownerId,$fileId,'$created')"
                        SqlQueryInsert -query $query
                    }
                }
            }
        }
        if (!$noSSN)
        {
            if ($hasSSNValue)
            {
                $query = "SELECT Id FROM $filesTableName WHERE FileName = '$FileName' AND Path = '$Path'"
                $sqlResult = SqlQueryReturn($query)
                foreach($row in $sqlResult)
                {
                    $fileId = $row.Id
                }

                foreach ($match in $result.SSNMatches)
                {
                    if ($match -ne $null -AND $match.Trim() -ne '')
                    {
                        $created = Get-Date
                        $query = "INSERT INTO ScanSSN (Match,OwnerId,FileId,Created) VALUES ('$match',$ownerId,$fileId,'$created')"
                        SqlQueryInsert -query $query
                    }
                }
            }
        }
        if (!$noCC)
        {
            if ($hasCCValue)
            {
                $query = "SELECT Id FROM $filesTableName WHERE FileName = '$FileName' AND Path = '$Path'"
                $sqlResult = SqlQueryReturn($query)
                foreach($row in $sqlResult)
                {
                    $fileId = $row.Id
                }

                foreach ($match in $result.CCMatches)
                {
                    if ($match -ne $null -AND $match.Trim() -ne '')
                    {
                        $created = Get-Date
                        $query = "INSERT INTO ScanCC (Match,OwnerId,FileId,Created) VALUES ('$match',$ownerId,$fileId,'$created')"
                        SqlQueryInsert -query $query
                    }
                }
            }
        }
		#New-Object -TypeName PsObject -Property @{FileName=$fileName;Message="success";ParentFolderCurrent=$parentFolderCurrent;Query=$query}
	}
	catch
	{
		$fileName = $file.Name
		$query = $query.Replace(',','_')
		$path = $path.Replace(',','_')
		$fileName = $fileName.Replace(',','_')
		$line = $_.InvocationInfo.ScriptLineNumber
		$message = $line.ToString() + ":" + $_.Exception.Message
		write-error $message
        New-Object -TypeName PsObject -Property @{FileName=$FileName;Message=$message;Path=$Path;Query=$query}
        $global:currentErrorCount++
	}
}

#usage .\crawl_v10.ps1 |  select-object FileName,Message,ParentFolderCurrent, Query  | Export-Csv ".\crawl_users_001.csv"


