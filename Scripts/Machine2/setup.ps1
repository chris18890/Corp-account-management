param([string]$Domain)
If (!$Domain) {
    $Domain="$env:userdomain"
}
If ($env:computername -ne "$Domain-DC2") {
    $Machine = "$Domain-DC2"
    rename-computer $Machine
}
Cscript c:\windows\system32\SCRegEdit.wsf /ar 0
Cscript c:\windows\system32\scregedit.wsf /cs 1
Cscript c:\windows\system32\scregedit.wsf /AU 4
$IPNet1 = "10.71.104"
$IPNetSite = $IPNet1
$IPAddress = "$IPNetSite.3"
$Netmask = "21"
$GatewayAddress = "$IPNet1.1"
$DNSAddress = @("$IPNet1.2", "$IPNet1.3")
Get-NetAdapter -Name Ethernet | Rename-NetAdapter -NewName LAN -PassThru
New-NetIPAddress -IPAddress $IPAddress -InterfaceAlias "LAN" -DefaultGateway $GatewayAddress -AddressFamily IPv4 -PrefixLength $Netmask
Set-DnsClientServerAddress -InterfaceAlias "LAN" -ServerAddresses $DNSAddress
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
