
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[System.Reflection.Assembly]::LoadWithPartialName("System.Xml")
[System.Reflection.Assembly]::LoadWithPartialName("System.Security")

Add-Type -Path 'c:\scripts\pdf\BouncyCastle.Crypto.dll'
Add-Type -Path 'c:\scripts\pdf\itextsharp.dll'


$filePath = 'C:\scripts\pdf\YTD - On Time Report.pdf'


$pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $filePath

$allText = ""

for($counter = 1; $counter -le $pdf.NumberOfPages; $counter++)
{
    $text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf, $counter)
    $allText = $allText + $text
    #$text
}

$allText

#write-host $pdf.NumberOfPages