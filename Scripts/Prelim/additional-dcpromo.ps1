#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

# Mode: Standalone / No Shared Modules

param([string]$Domain,[string]$DomainSuffix)

If (!$Domain) {
    $Domain = "$env:userdomain"
}
If (!$DomainSuffix) {
    $DomainSuffix = READ-HOST 'Enter a public FQDN- '
}
$ComputerName = "$env:computername"
if ((Get-CimInstance win32_computersystem).partofdomain -eq $false) {
    Add-Computer -DomainName "$Domain.$DomainSuffix" -Restart
} else {
    Install-WindowsFeature -name AD-Domain-Services, FS-DFS-Namespace, FS-DFS-Replication -IncludeManagementTools -ComputerName $ComputerName
    Install-ADDSDomainController -Credential (Get-Credential) -DomainName "$Domain.$DomainSuffix" -InstallDns:$true -NoGlobalCatalog:$false -Force:$true
}
