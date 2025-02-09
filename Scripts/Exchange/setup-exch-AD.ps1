$Domain = "$env:userdomain"
$Member = "$env:username"
msiexec.exe /i "\\$Domain\Store\Software\rewrite_amd64_en-US.msi" /quiet
Start-Process -Filepath "\\$Domain\Store\Software\vcredist_x64_2013.exe" -Argumentlist "/Q" -wait
.\UCMARedist\Setup.exe -q
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAD /OrganizationName:"$Domain"
Remove-ADGroupMember -Identity "Enterprise Admins" -Members $Member -Confirm:$False
Remove-ADGroupMember -Identity "Schema Admins" -Members $Member -Confirm:$False
Restart-Computer
