

$global:unixEpoch = Get-Date -Date "01/01/1970"
$global:DataSource = $PSScriptRoot + "\FileToOneDrive.db"

#function InitCrawl($ownerId, $email, $startPath, $doConvert)
function InitCrawl($ownerId, $startPath, $doConvert)
{
	$global:word = $null
	$global:excel = $null
	$global:powerpoint = $null

	$filesUsersTableName = ""
	$filesTableName = ""
	if ($doConvert -eq $true)
	{
		$filesUsersTableName = "Files_Users_Conversion"
		$filesTableName = "Files_OneDrive_Conversion"
	}
	else
	{
		$filesUsersTableName = "Files_Users"
		$filesTableName = "Files_OneDrive"
	}


    $StartPath = $path
    $ownerId = $ownerId
    $global:currentFileCount = 0
    $global:currentFileSize = 0
    $global:currentErrorCount = 0
    
    CrawlFolder $StartPath $ownerId
	
    InsertUserEntry $ownerId $global:currentFileCount $global:currentFileSize $global:currentErrorCount
    UpdateExtensions $ownerId
    UpdateFileTotals $ownerId
    UpdateOfficeErrorTotals $ownerId
    UpdateMacroCount $ownerId
    UpdateOldOfficeCount $ownerId
    UpdatePathLengthCount  $ownerId
    UpdateNoAccessErrorTotals $ownerId
	
    if ($doConvert -eq $true)
    {
        UpdateOfficeConversion $ownerId
    }
    #clean up orphaned office instances
    Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "POWERPNT" -Force -ErrorAction SilentlyContinue
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

function CrawlFolder($path, $ownerId, $currentDepth)
{
    try
    {
	    foreach ($file in Get-ChildItem -LiteralPath $path -File -ErrorAction Stop)
	    {
		    $tempPath = $path
		    write-host $file.Name 
		    InsertRow $file $tempPath $ownerId $currentDepth
	    }
        foreach ($folder in Get-ChildItem -Path $path -Directory -ErrorAction Stop)
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

#function InsertUserEntry($email, $ownerId, $fileCount, $fileSize, $errorCount)
function InsertUserEntry($ownerId, $fileCount, $fileSize, $errorCount)
{
    $query = ''
    try
    {
		$today = Get-Date
		$created = [int] (New-TimeSpan -Start $unixEpoch -End $today).TotalSeconds
        #$query = "Insert INTO  $filesUsersTableName  (Email, OwnerId,FileCountDisk,FileSizeDisk,ErrorCount,CreatedDate) Values('$email', $ownerId,$fileCount,$fileSize,$errorCount,$created)"
        $query = "Insert INTO  $filesUsersTableName  (OwnerId,FileCountDisk,FileSizeDisk,ErrorCount,CreatedDate) Values($ownerId,$fileCount,$fileSize,$errorCount,$created)"
        #write-host "Query:" $query -ForegroundColor Green 	   
	    Invoke-SqliteQuery -Query $Query -DataSource $global:DataSource
    }
	catch
	{
	    $line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
        New-Object -TypeName PsObject -Property @{FileName="InsertUserEntry";Message=$message;Path="";Query=""}
        $global:currentErrorCount++
	}
}


function ConvertDocument($path, $file, $saveAs)
{
    #################
   
    #################
    $result = New-Object -TypeName psobject 
    $result | Add-Member -MemberType NoteProperty -Name HasMacro -Value $false
    $result | Add-Member -MemberType NoteProperty -Name ConvertMessage -Value ""
    $result | Add-Member -MemberType NoteProperty -Name ConvertSuccess -Value $false

    $filePath = $path + "\" + $file.Name
    $name = $file.Name
    write-host "ConvertDocument()" $filePath 
	$parts = $filePath.Split('.')
	$extension = $parts[$parts.length - 1]
	$baseFileName = $filePath.Replace("." + $extension, "")
	$converted = $false
	$message = ""
    $oldFormat = $false
	try
	{
		if ($extension -eq "doc")
		{
            $oldFormat = $true

            if ($global:word -eq $null -OR $global:word.documents -eq $null)
            {
                 [gc]::collect()
                 [gc]::WaitForPendingFinalizers()
                 $global:word = new-object -comobject word.application
                 $global:word.Visible = $False
                 $global:word.DisplayAlerts = [Enum]::Parse([Microsoft.Office.Interop.Word.WdAlertLevel],"wdAlertsNone")
                 #new 7/27/19
                 #$global:excel.DisplayAlerts = $False;
                 $global:wordSaveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat],"wdFormatDocumentDefault")
                 $global:word.AutomationSecurity = 'msoAutomationSecurityForceDisable'
            }
            
            $testFilePath = $filePath + "x"
            if ([System.IO.File]::Exists($testFilePath) -eq $false)
            {
               
                #$savename = $filePath.ToLower() -replace ".doc", ".docx"
                $savename = $filePath.ToLower() + 'x'
                #copy to local location
                Write-Host "opening:" $filePath  
                write-host "Saving as :" $savename -ForegroundColor Cyan
                try
                {			    
                    #$opendoc = $global:word.documents.open($filePath,$false,$true)
                    #new 7/27/19
                    $opendoc = $global:word.documents.OpenNoRepairDialog($filePath,$false,$true)
                   <#
                        $openCounter = 0
                        while (-not $global:word.Ready) 
                        {
                         write-host '.' -NoNewline
                          sleep 1
                          $openCounter++
                          if ($openCounter -gt 10)
                         {
                            $global:word = $null
                            Stop-Process -Name "WINWORD" -Force
                            throw "file took too long to open"
                         }
                        }
                        #>
                        if ($opendoc -eq $null)
                        {
                            write-host "DOC IS NULL" -ForegroundColor yellow
                        }
              		if ($saveAs -eq $true)
                    {
                        $opendoc.saveas([ref]"$savename", [ref]$global:wordSaveFormat);
                    }
			        $opendoc.close($false);
			        $converted = $true
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
                        $opendoc = $global:word.documents.OpenNoRepairDialog($tempFilePath,$false,$true)
                        
                        <#
                        $openCounter = 0
                        while (-not $global:word.Ready) 
                        {
                         write-host '.' -NoNewline
                          sleep 1
                          $openCounter++
                          if ($openCounter -gt 10)
                         {
                            $global:word = $null
                            Stop-Process -Name "WINWORD" -Force
                            throw "file took too long to open"
                         }
                        }
                        #>
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
            }
            else
            {
                 $oldFormat = $false
            }
		}
		elseif ($extension -eq "xls")
		{
            $oldFormat = $true
            if ($global:excel -eq $null -OR $global:excel.workbooks -eq $null)
            {
                [gc]::collect()
                [gc]::WaitForPendingFinalizers()
                $global:excel = new-object -comobject excel.application
                $global:excel.Visible = $False
                $global:excelSaveFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
                $global:excel.DisplayAlerts = $False;
                $global:excel.AutomationSecurity = 'msoAutomationSecurityForceDisable'
            }
            $testFilePath = $filePath + "x"
            if ([System.IO.File]::Exists($testFilePath) -eq $false)
            {
                try
                {
                   
                    $savename = $filePath.ToLower() + 'x'

			        $workBook  =  $global:excel.workbooks.open("$filePath", 0, 0, 5, "")
                    if ($workbook.HasVBProject)
                    {
                        $result.HasMacro = $true
                        $savename = $savename -Replace ".xlsx", ".xlsm"
                        $workBook.saveas([ref]"$savename", [ref][Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled);
                    }
                    else
                    {
              	      #$savename = ($filePath).substring(0,($filePath).lastindexOf("."))

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
                        $workBook  =  $global:excel.workbooks.open("$tempFilePath", 0, 0, 5, "")
                        if ($workbook.HasVBProject)
                        {
                             $result.HasMacro = $true
                             $tempSaveName = $tempSaveName -Replace ".xlsx", ".xlsm"
                             $workBook.saveas([ref]"$tempSaveName", [ref][Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled);
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
            }
            else
            {
                $oldFormat = $false
            }
		}
		elseif ($extension -eq "ppt")
		{
             $oldFormat = $true
             if ($global:powerpoint -eq $null -OR $global:powerpoint.Presentations -eq $null )
             {
                [gc]::collect()
                [gc]::WaitForPendingFinalizers()
                $global:powerpoint = New-Object -ComObject PowerPoint.application
                $global:powerpointSaveFormat = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentation 
                #$global:powerpoint.DisplayAlerts = $False;
                #$global:powerpoint.DisplayAlerts = [Enum]::Parse([Microsoft.Office.Interop.PowerPoint.WdAlertLevel],"wdAlertsNone")
                  $global:powerpoint.DisplayAlerts =  [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone
                $global:powerpoint.AutomationSecurity = 'msoAutomationSecurityForceDisable'

             }
            $testFilePath = $filePath + "x"
            if ([System.IO.File]::Exists($testFilePath) -eq $false)
            {
                try
                {
			        $presentation = $global:powerpoint.Presentations.open($filePath, $true, $null, $false)
			        #$savename = ($filePath).substring(0,($filePath).lastindexOf("."))
                    $savename = $filePath.ToLower() + 'x'

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
                        $presentation = $global:powerpoint.Presentations.open($tempFilePath, $true, $null, $false)
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
            }
            else
            {
                $oldFormat = $false
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
    return $result
}

function UpdateExtensions($ownerId)
{
    try
    {
        $query = "select Extension, count(*) as count from $filesTableName Where OwnerId = $ownerId AND Extension IS NOT NULL GROUP BY Extension"
		$Result = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
 
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
        $rowsAffected = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
    }
	catch
	{
    	$line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
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
        $Result = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
        $officeErrorCount = 0
        Foreach ($row in $Result) 
        {
            $officeErrorCount = $row.count
        }
        $query = "UPDATE  $filesUsersTableName  SET NoAccessCount = $officeErrorCount WHERE OwnerId = $ownerId"
	    $rowsAffected = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
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
        $query = "select count(*) as count from $filesTableName where Error IS NOT NULL AND Error <> '' AND Error <> 'No Access' AND OwnerId = $ownerId"
        $Result = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
     
        Foreach ($row in $Result) 
        {
            $officeErrorCount = $row.count
        }
        $query = "UPDATE  $filesUsersTableName  SET OfficeErrorCount = $officeErrorCount WHERE OwnerId = $ownerId"
        $rowsAffected = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
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
		$Result = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
        $macroCount = 0
        Foreach ($row in $Result) 
        {
            $macroCount = $row.count
        }
        $query = "UPDATE  $filesUsersTableName  SET MacroCount = $macroCount WHERE OwnerId = $ownerId"
        $rowsAffected = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
	 }
	catch
	{
        $line = $_.InvocationInfo.ScriptLineNumber
		$message = $line + " " + $_.Exception.Message
		write-error $message
        $global:currentErrorCount++
        New-Object -TypeName PsObject -Property @{FileName="UpdatMacroCount";Message=$message;Path="";Query=""}

	}

}

function UpdateOldOfficeCount($ownerId)
{
  try
    {
        $query = "select count(*) as count from $filesTableName where (Extension = 'doc' OR Extension = 'xls' OR Extension = 'ppt') AND OwnerId = $ownerId"
       	$Result = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
        $oldOfficeCount = 0
        Foreach ($row in $Result) 
        {
            $oldOfficeCount = $row.count
        }
        $query = "UPDATE  $filesUsersTableName  SET OldOfficeCount = $oldOfficeCount WHERE OwnerId = $ownerId"
        $rowsAffected = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
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
        $Result = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
        $pathLengthCount = 0
        Foreach ($row in $Result) 
        {
            $pathLengthCount = $row.count
        }
        $query = "UPDATE  $filesUsersTableName  SET PathLengthCount = $pathLengthCount WHERE OwnerId = $ownerId"
        $rowsAffected = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
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
        $Result = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
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
        $rowsAffected = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
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
        $Result = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
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
		$rowsAffected = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
}

function InsertRow($file, $path, $ownerId, $currentDepth)
{
<#
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
         $result | Add-Member -MemberType NoteProperty -Name ConvertMessage -Value $line.ToString() + ":" + $_.Exception.Message
         $result | Add-Member -MemberType NoteProperty -Name ConvertSuccess -Value $false

    }
    $path = $path -Replace "'", "''"
    $hasMacro = $false
    $tempFilePath = $path + "\" + $file.Name
    $tempFilePathLength = $tempFilePath.Length

    $convertSuccess = $result.ConvertSuccess
    $convertMessage = $result.ConvertMessage
    $hasMacroValue = 0
    $convertSuccessValue = 0

    if ($result.HasMacro -eq $true)
    {
        $hasMacroValue = 1
    }
    if ($result.ConvertSuccess -eq $true)
    {
        $convertSuccessValue = 1
    }
	#>
	$tempFilePath = $path + "\" + $file.Name
    $tempFilePathLength = $tempFilePath.Length
	
    $global:currentFileSize +=  ($file.length / 1000000.0)
    $global:currentFileCount += 1
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
		$Size = $file.length / 1000000.0

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
		
		
		$RelativeFolder = $tempParentFolderCurrent.Replace($StartPath.ToLower(), "")
		#$query = "INSERT INTO $filesTableName (FileName, Created, Modified, Author, Extension, Size, OwnerId, Ignore, Path, FolderDepth, ParentFolder, RelativeFolder,OfficeOpen,Error,PathLength, HasMacro,"
		#$query += "Folder01, Folder02, Folder03, Folder04, Folder05, Folder06, Folder07, Folder08, Folder09, Folder10, Folder11, Folder12, Folder13, Folder14, Folder15, Folder16, Folder17, Folder18, Folder19, Folder20) "
		#$query += " VALUES ('$FileName', '$createdSeconds', '$modifiedSeconds', '$Author', '$Extension', '$Size', $ownerId, '$Ignore', '$Path', '$FolderDepth','$tempParentFolderCurrent', '$RelativeFolder',$convertSuccessValue,'$convertMessage',$tempFilePathLength,$hasMacroValue,"
		#$query += " '$Folder01', '$Folder02', '$Folder03', '$Folder04', '$Folder05', '$Folder06', '$Folder07', '$Folder08', '$Folder09', '$Folder10', '$Folder11', '$Folder12', '$Folder13', '$Folder14', '$Folder15', '$Folder16', '$Folder17', '$Folder18', '$Folder19', '$Folder20')"

		$query = "INSERT INTO $filesTableName (FileName, Created, Modified, Author, Extension, Size, OwnerId, Ignore, Path, FolderDepth, ParentFolder, RelativeFolder,PathLength,"
		$query += "Folder01, Folder02, Folder03, Folder04, Folder05, Folder06, Folder07, Folder08, Folder09, Folder10, Folder11, Folder12, Folder13, Folder14, Folder15, Folder16, Folder17, Folder18, Folder19, Folder20) "
		$query += " VALUES ('$FileName', '$createdSeconds', '$modifiedSeconds', '$Author', '$Extension', '$Size', $ownerId, '$Ignore', '$Path', '$FolderDepth','$tempParentFolderCurrent', '$RelativeFolder',$tempFilePathLength,"
		$query += "'$Folder01', '$Folder02', '$Folder03', '$Folder04', '$Folder05', '$Folder06', '$Folder07', '$Folder08', '$Folder09', '$Folder10', '$Folder11', '$Folder12', '$Folder13', '$Folder14', '$Folder15', '$Folder16', '$Folder17', '$Folder18', '$Folder19', '$Folder20')"
		write-host $query -ForegroundColor Green
		$rowsAffected = Invoke-SqliteQuery -Query $query -DataSource $global:DataSource
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


