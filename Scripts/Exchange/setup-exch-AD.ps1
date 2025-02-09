$Domain = "$env:userdomain"
$Member = "$env:username"
.\UCMARedist\Setup.exe -q
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAD /OrganizationName:"$Domain"
msiexec.exe /i "\\$Domain\Store\Software\rewrite_amd64_en-US.msi" /quiet
\\$Domain\Store\Software\vcredist_x64_2013.exe /install /quiet /norestart
Remove-ADGroupMember -Identity "Enterprise Admins" -Members $Member -Confirm:$False
Remove-ADGroupMember -Identity "Schema Admins" -Members $Member -Confirm:$False
Restart-Computer
