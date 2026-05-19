#Requires -RunAsAdministrator

# Execution Tier: Tier-0
# Mode: Standalone / No Shared Modules

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Domain
    ,[Parameter(Mandatory)][ValidateSet('Azure','Local')][string]$Platform
    ,[Parameter(Mandatory)][ValidateSet('DC1','RTR','DC2','DC3','RootCA','InterCA','WSUS','EXCH')][string]$Role
    ,[string]$EnvironmentConfig = (Join-Path (Split-Path $PSScriptRoot -Parent) 'environment.psd1')
)

Set-StrictMode -Version Latest

if (-not (Test-Path -LiteralPath $EnvironmentConfig)) {
    throw "environment.psd1 not found at '$EnvironmentConfig'"
}
$Env = Import-PowerShellDataFile -LiteralPath $EnvironmentConfig

$IPNet1   = $Env.Network.SiteSubnets[0]
$IPNet2   = $Env.Network.SiteSubnets[1]
$Netmask  = $Env.Network.VnetPrefix
$Hosts    = $Env.Network.HostOffsetsLocal
$GatewayAddress = "$IPNet1.$($Hosts.RTR)"

switch ($Role) {
    "DC1" {
        If ($env:computername -ne "$Domain-$Role") {
            $Machine = "$Domain-$Role"
            rename-computer $Machine
        }
        $IPAddress = "$IPNet1.$($Hosts.DC1)"
        $DNSAddress = @("$IPNet1.$($Hosts.DC1)", "$IPNet1.$($Hosts.DC2)")
    }
    "RTR" {
        If ($env:computername -ne "$Domain-$Role") {
            $Machine = "$Domain-$Role"
            rename-computer $Machine
        }
        $IPAddress = "$IPNet1.$($Hosts.RTR)"
        $DNSAddress = @("$IPNet1.$($Hosts.DC1)")
    }
    "DC2" {
        If ($env:computername -ne "$Domain-$Role") {
            $Machine = "$Domain-$Role"
            rename-computer $Machine
        }
        $IPAddress = "$IPNet1.$($Hosts.DC2)"
        $DNSAddress = @("$IPNet1.$($Hosts.DC1)", "$IPNet1.$($Hosts.DC2)")
    }
    "RootCA" {
        If ($env:computername -ne "$Domain-$Role") {
            $Machine = "$Domain-$Role"
            rename-computer $Machine
        }
        $IPAddress = "$IPNet1.$($Hosts.RootCA)"
        $DNSAddress = @("$IPNet1.$($Hosts.DC1)", "$IPNet1.$($Hosts.DC2)")
    }
    "InterCA" {
        If ($env:computername -ne "$Domain-$Role") {
            $Machine = "$Domain-$Role"
            rename-computer $Machine
        }
        $IPAddress = "$IPNet1.$($Hosts.InterCA)"
        $DNSAddress = @("$IPNet1.$($Hosts.DC1)", "$IPNet1.$($Hosts.DC2)")
    }
    "WSUS" {
        If ($env:computername -ne "$Domain-$Role") {
            $Machine = "$Domain-$Role"
            rename-computer $Machine
        }
        $IPAddress = "$IPNet1.$($Hosts.WSUS)"
        $DNSAddress = @("$IPNet1.$($Hosts.DC1)", "$IPNet1.$($Hosts.DC2)")
    }
    "EXCH" {
        If ($env:computername -ne "$Domain-$Role") {
            $Machine = "$Domain-$Role"
            rename-computer $Machine
        }
        $IPAddress = "$IPNet1.$($Hosts.EXCH)"
        $DNSAddress = @("$IPNet1.$($Hosts.DC1)", "$IPNet1.$($Hosts.DC2)")
    }
    "DC3" {
        If ($env:computername -ne "$Domain-$Role") {
            $Machine = "$Domain-$Role"
            rename-computer $Machine
        }
        $IPAddress = "$IPNet2.$($Hosts.DC1)"
        $DNSAddress = @("$IPNet1.$($Hosts.DC1)", "$IPNet1.$($Hosts.DC2)")
    }
}
if ($Platform -eq "Local") {
    $InternalInterface = $Env.Network.Interfaces.Internal
    Get-NetAdapter -Name Ethernet | Rename-NetAdapter -NewName $InternalInterface -PassThru
    New-NetIPAddress -IPAddress $IPAddress -InterfaceAlias $InternalInterface -DefaultGateway $GatewayAddress -AddressFamily IPv4 -PrefixLength $Netmask
    Set-DnsClientServerAddress -InterfaceAlias $InternalInterface -ServerAddresses $DNSAddress
}
Cscript c:\windows\system32\SCRegEdit.wsf /ar 0
Cscript c:\windows\system32\scregedit.wsf /cs 1
Cscript c:\windows\system32\scregedit.wsf /AU 4
Import-Module NetSecurity
Set-NetFirewallRule -DisplayName "File and Printer Sharing (Echo Request - ICMPv4-In)" -enabled True
Set-NetFirewallRule -DisplayName "File and Printer Sharing (Echo Request - ICMPv4-Out)" -enabled True
Set-NetFirewallRule -DisplayName "File and Printer Sharing (SMB-In)" -enabled True
Set-NetFirewallRule -DisplayName "File and Printer Sharing (SMB-Out)" -enabled True
Set-NetFirewallRule -DisplayGroup "Windows Remote Management" -enabled True
Set-NetFirewallRule -DisplayGroup "Remote Volume Management" -enabled True
Set-NetFirewallRule -DisplayGroup "Remote Event Log Management" -enabled True
Set-NetFirewallRule -DisplayGroup "Remote Scheduled Tasks Management" -enabled True
Restart-Computer
