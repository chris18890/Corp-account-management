#Requires -RunAsAdministrator

# Execution Tier: Tier-0
# Mode: Standalone / No Shared Modules

param(
    [Parameter(Mandatory)][string]$Domain
    ,[Parameter(Mandatory)][string]$DomainSuffix
)

Set-StrictMode -Version Latest

$ComputerName = $env:computername
Install-WindowsFeature -name AD-Domain-Services, FS-DFS-Namespace, FS-DFS-Replication -IncludeManagementTools -ComputerName "$ComputerName"
Install-ADDSForest -DomainName "$Domain.$DomainSuffix" -ForestMode "WinThreshold" -Force:$true
