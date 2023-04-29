$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$ParentOU = "Domain Computers"
$Location = "OU=$ParentOU,$EndPath"
$GPOLocation = "c:\scripts\prelim\gpos"
$GPOName = "LAPS"

Update-LapsADSchema
Set-LapsADComputerSelfPermission -Identity "$ParentOU"
Set-LapsADReadPasswordPermission -Identity "Desktops" -AllowedPrincipals "$Domain\RG_Desktop_Admins"
Set-LapsADResetPasswordPermission -Identity "Desktops" -AllowedPrincipals "$Domain\RG_Desktop_Admins"
Set-LapsADReadPasswordPermission -Identity "Laptops" -AllowedPrincipals "$Domain\RG_Desktop_Admins"
Set-LapsADResetPasswordPermission -Identity "Laptops" -AllowedPrincipals "$Domain\RG_Desktop_Admins"
Set-LapsADReadPasswordPermission -Identity "Servers" -AllowedPrincipals "$Domain\RG_Server_Admins"
Set-LapsADResetPasswordPermission -Identity "Servers" -AllowedPrincipals "$Domain\RG_Server_Admins"
Import-GPO -BackupGpoName $GPOName -TargetName $GPOName -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
New-GPLink -name $GPOName -target $Location -LinkEnabled Yes -enforced yes -ErrorAction Stop
