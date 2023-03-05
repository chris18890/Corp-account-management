$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$ParentOU = "Domain Computers"
$Location = "OU=$ParentOU,$EndPath"
$GPOLocation = "c:\scripts\prelim\gpos"
$GPOName = "LAPS"

Import-Module AdmPwd.ps
Update-AdmPwdADSchema
Set-AdmPwdComputerSelfPermission -OrgUnit "$ParentOU"
Import-GPO -BackupGpoName $GPOName -TargetName $GPOName -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
New-GPLink -name $GPOName -target $Location -LinkEnabled Yes -enforced yes -ErrorAction Stop
