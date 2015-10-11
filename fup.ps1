# This copies follow-up patient data
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$dte=Get-Date -Format "M/d/yyyy"
$nmr=[System.Windows.Forms.Clipboard]::GetText()
$fin=$nmr.Split(",")
$ret=$fin[0] + "`t" + $fin[1] + "`t" + $dte
[System.Windows.Forms.Clipboard]::SetText($ret)