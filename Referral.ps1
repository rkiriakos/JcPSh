# This creates new patient referral comment
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$dte=[datetime] ([System.Windows.Forms.Clipboard]::GetText())
$ts=New-TimeSpan -Start $dte
$dd=$dte.ToShortDateString()
$dl=$ts.Days
[Windows.Forms.Clipboard]::SetText("Referral $dd delay $dl days")