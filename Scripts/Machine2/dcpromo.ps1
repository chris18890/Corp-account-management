param([string]$Domain)
If (!$Domain) {
    $Domain = $env:userdomain
}
$ComputerName = $env:computername
if ((gwmi win32_computersystem).partofdomain -eq $false) {
    Add-Computer -DomainName "$Domain" -Restart
}
Install-WindowsFeature -name AD-Domain-Services, FS-DFS-Namespace, FS-DFS-Replication -IncludeManagementTools -ComputerName $ComputerName
Install-ADDSDomainController -Credential (Get-Credential) -DomainName "$Domain" -InstallDns:$true -NoGlobalCatalog:$false -Force:$true
