param([string]$Domain,[string]$DomainSuffix)
If (!$Domain) {
    $Domain = READ-HOST 'Enter a NETBIOS Domain name- '
}
$ServerName = "$env:computername"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$ParentOU = "Domain Computers"
$SID500 = "$env:username"
$Location = "OU=$ParentOU,$EndPath"
$CDIR = "24"
$IPNet = @("10.71.104", "10.71.105")
$Site = @("Site1", "Site2")
$FeatureName = @("rsat-ad-powershell", "rsat-adds", "GPMC", "RSAT-DFS-Mgmt-Con", "rsat-dns-server", "rsat-dhcp", "WDS-AdminPack", "RSAT-ADCS", "UpdateServices-RSAT")
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
    If (!$DomainSuffix) {
        $DomainSuffix = READ-HOST 'Enter a public FQDN- '
    }
    $DNSSuffix = "$Domain.$DomainSuffix"
    Add-Computer -DomainName "$DNSSuffix" -Restart
} else {
    if ((gwmi win32_computersystem).partofdomain -eq $true) {
        Get-ADComputer $ServerName | Move-ADObject -TargetPath "ou=Servers,$Location"
        For ($i=0; $i -le $IPNet.Length -1; $i++){
            $ScopeNetwork = $IPNet[$i] + ".0"
            $ScopeName = "$ScopeNetwork/$CDIR"
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
        Remove-ADGroupMember -Identity "Enterprise Admins" -Members $SID500 -Confirm:$False
    }
}
