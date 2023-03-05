param([string]$Domain)
If (!$Domain) {
    $Domain = READ-HOST 'Enter a NETBIOS Domain name- '
}
$ServerName = "$env:computername"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$ParentOU = "Domain Computers"
$Location = "OU=$ParentOU,$EndPath"
$Netmask = "255.255.255.0"
$CDIR = "24"
$Router = "1"
$DNSServer2 = "2"
$DNSServer3 = "3"
$StartRange = "10"
$EndRange = "254"
$IPNet = @("10.71.104", "10.71.105")
$Site = @("Site1", "Site2")
$FeatureName = @("rsat-ad-powershell", "rsat-adds", "GPMC", "routing", "dhcp", "RSAT-DFS-Mgmt-Con", "rsat-dns-server", "WDS-AdminPack", "RSAT-ADCS", "UpdateServices-RSAT")
$IPAddress1 = $IPNet[0] + "." + $Router
$IPAddress2 = $IPNet[0] + "." + $DNSServer2
$IPAddress3 = $IPNet[0] + "." + $DNSServer3
$ExternalInterface = "WAN"
$InternalInterface = "LAN"
Foreach ($Feature in $FeatureName){
    if (Get-WindowsFeature -Name $Feature | Where InstallState -Eq Installed) {
        Write-Host $Feature "is already installed" -ForegroundColor Green
    } else {
        Write-Host "installing" $Feature
        install-windowsfeature -IncludeManagementTools $Feature
        Write-Host "installed" $Feature
    }
}
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
if ((gwmi win32_computersystem).partofdomain -eq $false) {
    Add-Computer -DomainName "$DNSSuffix" -Restart
} else {
    if ((gwmi win32_computersystem).partofdomain -eq $true) {
        Set-DhcpServerv4Binding -BindingState $true -InterfaceAlias "$InternalInterface"
        Add-DhcpServerSecurityGroup -ComputerName $ServerName
        Restart-service dhcpserver
        Add-DhcpServerInDC "$ServerName" "$IPAddress1"
        try {
            Add-LocalGroupMember -Group "DHCP Administrators" -Member "$Domain\RG_Server_Admins" -ErrorAction Stop
        } catch [Microsoft.PowerShell.Commands.MemberExistsException] {
            Write-Host "$Domain\RG_Server_Admins is already a member of DHCP Administrators" -ForegroundColor Green
        }
        Set-ItemProperty -Path registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ServerManager\Roles\12 -Name ConfigurationState -Value 2
        Get-ADComputer $ServerName | Move-ADObject -TargetPath "ou=Servers,$Location"
        Set-DhcpServerv4OptionValue -Router "$IPAddress1"
        Set-DhcpServerv4OptionValue -OptionId 42 -value "$IPAddress2"
        Set-DhcpServerv4OptionValue -DnsServer "$IPAddress2", "$IPAddress3" -Force
        Set-DhcpServerv4OptionValue -DnsDomain "$DNSSuffix"
        For ($i=0; $i -le $IPNet.Length -1; $i++){
            $ScopeNetwork = $IPNet[$i] + ".0"
            $ScopeName = "$ScopeNetwork/$CDIR"
            $ScopeStart = $IPNet[$i] + ".$StartRange"
            $ScopeEnd = $IPNet[$i] + ".$EndRange"
            $ScopeNetmask = $Netmask
            $ScopeDNS2 = $IPNet[$i] + ".$DNSServer2"
            $ScopeDNS3 = $IPNet[$i] + ".$DNSServer3"
            $ScopeRouter = $IPNet[$i] + ".$Router"
            $Error.Clear()
            try {
                Add-DnsServerPrimaryZone -NetworkID "$ScopeName" -ReplicationScope "Domain" -DynamicUpdate:Secure -ComputerName "$DNSSuffix" -ErrorAction Stop
            } catch [Microsoft.Management.Infrastructure.CimException] {
                switch ($Error[0].Exception.ErrorCode) {
                    "Win32 9609"{ # 'The specified object already exists'
                        Write-Host "Zone" $ScopeName "already exists" -ForegroundColor Green
                    }
                    default {
                        Write-Host "ERROR: An unexpected error occurred while attempting to create Zone" $ScopeName":`n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                    }
                }
            }
            try {
                Add-DhcpServerv4Scope -Name $Site[$i] -StartRange "$ScopeStart" -EndRange "$ScopeEnd" -SubnetMask "$ScopeNetmask" -State Active -ErrorAction Stop
            } catch [Microsoft.Management.Infrastructure.CimException] {
                switch ($Error[0].Exception.ErrorCode) {
                    "DHCP 20052" { # 'The specified object already exists'
                        Write-Host "Scope "$Site[$i]" already exists" -ForegroundColor Green
                    }
                    default {
                        Write-Host "ERROR: An unexpected error occurred while attempting to create Scope" $Site[$i]":`n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                    }
                }
            }
            Set-DhcpServerv4OptionValue -ScopeId "$ScopeNetwork" -DnsServer "$ScopeDNS2", "$ScopeDNS3" -Force -Router "$ScopeRouter"
            try {
                New-ADReplicationSite -Name $Site[$i]
            } catch [Microsoft.ActiveDirectory.Management.ADException] {
                switch ($Error[0].Exception.ErrorCode) {
                    8305 { # 'The specified object already exists'
                        Write-Host "Site" $Site[$i] "already exists" -ForegroundColor Green
                    }
                    default {
                        Write-Host "ERROR: An unexpected error occurred while attempting to create Site" $Site[$i]":`n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                    }
                }
            }
            try {
                New-ADReplicationSubnet -Name "$ScopeName" -Site $Site[$i]
            } catch [Microsoft.ActiveDirectory.Management.ADException] {
                switch ($Error[0].Exception.ErrorCode) {
                    8305 { # 'The specified object already exists'
                        Write-Host "Subnet" $ScopeName "already exists" -ForegroundColor Green
                    }
                    default {
                        Write-Host "ERROR: An unexpected error occurred while attempting to create Subnet" $ScopeName":`n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                    }
                }
            }
            Set-ADReplicationSiteLink -Identity "DEFAULTIPSITELINK" -SitesIncluded @{Add=$Site[$i]}
        }
        Set-ADReplicationSiteLink -Identity "DEFAULTIPSITELINK" -Cost "10" -ReplicationFrequencyInMinutes "15"
        Move-ADDirectoryServer -Identity "$Domain-DC1" -Site $Site[0]
        Install-RemoteAccess -VpnType RoutingOnly -Legacy
        cmd.exe /c "netsh routing ip nat install"
        cmd.exe /c "netsh routing ip nat add interface $ExternalInterface"
        cmd.exe /c "netsh routing ip nat set interface $ExternalInterface mode=full"
        cmd.exe /c "netsh routing ip nat add interface $InternalInterface"
    }
}
