#Requires -RunAsAdministrator

# Execution Tier: Tier-0
# Mode: Standalone / No Shared Modules

param(
    [Parameter(Mandatory)][string]$Domain
    ,[Parameter(Mandatory)][string]$DomainSuffix
)

Set-StrictMode -Version Latest

$ComputerName = "$env:computername"
if ((Get-CimInstance win32_computersystem).partofdomain -eq $false) {
    Add-Computer -DomainName "$Domain.$DomainSuffix" -Restart
} else {
    $DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
    $requiredGroups = @('Domain Admins')
    $groups = $requiredGroups | ForEach-Object {
        Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
    }
    if (-not $groups) {
        Write-Host "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
        throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
    }
    Install-WindowsFeature -name AD-Domain-Services, FS-DFS-Namespace, FS-DFS-Replication -IncludeManagementTools -ComputerName $ComputerName
    Install-ADDSDomainController -Credential (Get-Credential) -DomainName "$Domain.$DomainSuffix" -InstallDns:$true -NoGlobalCatalog:$false -Force:$true
}
