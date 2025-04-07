param([string]$Domain,[string]$DomainSuffix)
If (!$Domain) {
    $Domain = READ-HOST 'Enter a NETBIOS Domain name- '
}
If (!$DomainSuffix) {
    $DomainSuffix = READ-HOST 'Enter a public FQDN- '
}
$DNSName = "$Domain.$DomainSuffix"
$ComputerName = $env:computername
Install-WindowsFeature -name AD-Domain-Services, FS-DFS-Namespace, FS-DFS-Replication -IncludeManagementTools -ComputerName "$ComputerName"
Install-ADDSForest -DomainName "$DNSName" -ForestMode "WinThreshold" -Force:$true
