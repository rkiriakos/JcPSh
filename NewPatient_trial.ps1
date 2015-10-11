[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$str = [string] ([System.Windows.Forms.Clipboard]::GetText())
$wd=New-Object -ComObject Word.Application
$doc=$wd.Documents.Open("C:\Users\royk\Documents\Card.docx")
$wd.Visible=$true
$fPath = $doc.AttachedTemplate.FullName
$cPath = $doc.AttachedTemplate.Path
$enc = New-Object -ComObject System.Security.Cryptography.MD5CryptoServiceProvider
$i=$str.Split(",")
$doc.Fields(1).Select
$wd.Selection.Text = $i.Item(0)

$doc.Fields(1).Select
$wd.Selection.Text = $i.Item(1)
