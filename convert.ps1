
 = 'C:\Users\cpreephim\Desktop\App Team\IQOQDQ\Data IQOQDQ\5XXXX_08_IQ_Format Parts_20XX-XX-XX_en.doc'
 = 'C:\Users\cpreephim\Desktop\App Team\IQOQDQ\Data IQOQDQ\5XXXX_08_IQ_Format Parts_20XX-XX-XX_en.docx'

 = New-Object -ComObject Word.Application
.Visible = False
.DisplayAlerts = 0

 = .Documents.Open()
.SaveAs2(, 16)
.Close()
.Quit()
Write-Host 'Converted successfully!'
