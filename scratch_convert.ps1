
$docPath = "C:\Users\cpreephim\Desktop\App Team\IQOQDQ\Data IQOQDQ\5XXXX_08_IQ_Format Parts_20XX-XX-XX_en.doc"
$docxPath = "C:\Users\cpreephim\Desktop\App Team\IQOQDQ\Data IQOQDQ\5XXXX_08_IQ_Format Parts_20XX-XX-XX_en.docx"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$doc = $word.Documents.Open($docPath)
$doc.SaveAs2($docxPath, 16)
$doc.Close([ref]$false)
$word.Quit()
Write-Host "Converted successfully to: $docxPath"
