#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

# Mode: Standalone / No Shared Modules

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Domain
    ,[Parameter(Mandatory)][ValidateSet('Azure','Local')][string]$Platform
    ,[string]$IPNet1 = "10.71.104"
    ,[string]$Netmask = "21"
)

If ($env:computername -ne "$Domain-DC2") {
    $Machine = "$Domain-DC2"
    rename-computer $Machine
}
If (!$IPNet1) {
    $IPNet1 = "10.71.104"
}
$IPNetSite = $IPNet1
$IPAddress = "$IPNetSite.3"
If (!$Netmask) {
    $Netmask = "21"
}
$GatewayAddress = "$IPNetSite.1"
$DNSAddress = @("$IPNetSite.2", "$IPNetSite.3")
if ($Platform -eq "Local") {
    Get-NetAdapter -Name Ethernet | Rename-NetAdapter -NewName LAN -PassThru
    New-NetIPAddress -IPAddress $IPAddress -InterfaceAlias "LAN" -DefaultGateway $GatewayAddress -AddressFamily IPv4 -PrefixLength $Netmask
    Set-DnsClientServerAddress -InterfaceAlias "LAN" -ServerAddresses $DNSAddress
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
