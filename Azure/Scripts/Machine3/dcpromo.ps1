param([string]$Domain,[string]$DomainSuffix)
If (!$Domain) {
    $Domain = $env:userdomain
}
$ComputerName = $env:computername
if ((gwmi win32_computersystem).partofdomain -eq $false) {
    If (!$DomainSuffix) {
        $DomainSuffix = READ-HOST 'Enter a public FQDN- '
    }
    Add-Computer -DomainName "$Domain.$DomainSuffix" -Restart
} else {
    Install-WindowsFeature -name AD-Domain-Services, FS-DFS-Namespace, FS-DFS-Replication -IncludeManagementTools -ComputerName $ComputerName
    Install-ADDSDomainController -Credential (Get-Credential) -DomainName "$Domain" -InstallDns:$true -NoGlobalCatalog:$false -Force:$true
}
