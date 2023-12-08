$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$ParentOU = "Domain Computers"
$Location = "OU=$ParentOU,$EndPath"
$GPOLocation = "c:\scripts\prelim\gpos"
$GPOName = "LAPS"

Update-LapsADSchema
Set-LapsADComputerSelfPermission -Identity "$Location"
Set-LapsADReadPasswordPermission -Identity "OU=Desktops,$Location" -AllowedPrincipals "$Domain\RG_Desktop_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=Desktops,$Location" -AllowedPrincipals "$Domain\RG_Desktop_Admins"
Set-LapsADReadPasswordPermission -Identity "OU=Laptops,$Location" -AllowedPrincipals "$Domain\RG_Desktop_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=Laptops,$Location" -AllowedPrincipals "$Domain\RG_Desktop_Admins"
Set-LapsADReadPasswordPermission -Identity "OU=Servers,$Location" -AllowedPrincipals "$Domain\RG_Server_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=Servers,$Location" -AllowedPrincipals "$Domain\RG_Server_Admins"
Import-GPO -BackupGpoName $GPOName -TargetName $GPOName -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
New-GPLink -name $GPOName -target $Location -LinkEnabled Yes -enforced yes -ErrorAction Stop
