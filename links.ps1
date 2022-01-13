$global:word = new-object -comobject word.application
$global:word.Visible = $False
$global:word.DisplayAlerts = [Enum]::Parse([Microsoft.Office.Interop.Word.WdAlertLevel],"wdAlertsNone")


$opendoc = $global:word.documents.OpenNoRepairDialog('c:\scripts\tempdocs\testlinks.docx',$false,$true,$false,'')
if ($opendoc.Hyperlinks.Count -gt 0)
{
    write-host 'FOUND WORD LINK'
}
$opendoc.Close($false)



$global:excel = new-object -comobject excel.application
$global:excel.Visible = $False
$workBook  =  $global:excel.workbooks.open('c:\scripts\tempdocs\testlinks.xlsx', $false, $true, 5, "")
foreach($worksheet in $workBook.worksheets)
{
    if ($worksheet.Hyperlinks.Count -gt 0)
    {
        write-host 'FOUND EXCEL LINK!!' -f WHITE
    }
}
$workBook.Close($false)



 $global:powerpoint = New-Object -ComObject PowerPoint.application
                    $global:powerpointSaveFormat = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentation 
                    $global:powerpoint.DisplayAlerts =  [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone
                    $global:powerpoint.AutomationSecurity = 'msoAutomationSecurityForceDisable'

        $presentation =             $global:powerpoint.Presentations.open('c:\scripts\tempdocs\testlinks.pptx', $true, $null, $false)
        
        foreach($slide in $presentation.Slides)
        {
            if ($slide.Hyperlinks.Count -gt 0)
            {
                write-host 'FOUND POWERPINT LINK'
            }
        }

        $presentation.Close()