[CmdletBinding()]
param(
    [string]$SubscriptionId
    , [string]$TenantID
    , [string]$MyIPAddress
    , [string]$NameStem
    , [string]$Owner
    , [string]$LogFile
    , [switch]$RemoveIncomplete
)

Set-StrictMode -Version Latest
Import-Module Az
Import-Module (Join-Path $PSScriptRoot 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$EnvConfig = Get-EnvironmentConfig

if (!$SubscriptionId) {
    $SubscriptionId = READ-HOST 'Enter Subscription ID - '
}
if (!$TenantID) {
    $TenantID = READ-HOST 'Enter Tenant ID - '
}
if (!$MyIPAddress) {
    $MyIPAddress = READ-HOST 'Enter IPv4 address as a CIDR /32 - '
}
if (!$NameStem) {
    $NameStem = READ-HOST 'Enter Name stem - '
}
if (!$Owner) {
    $Owner = READ-HOST 'Enter Owner Name - '
}
$IPNet         = $EnvConfig.Network.SiteSubnets
$Site          = $EnvConfig.Network.SiteNames
$Subnet        = @()
$SubnetCDIR    = $EnvConfig.Network.SiteCidr
$VnetCDIR      = $EnvConfig.Network.VnetPrefix
$VnetDNSServer = "$($IPNet[0]).$($EnvConfig.Network.HostOffsetsAzure.DC1)"
$VnetPrefix    = "$($IPNet[0]).0/$VnetCDIR"
$ResourceGroupName  = "${NameStem}_Infra_RG"
$Location           = $EnvConfig.Azure.Location
$VirtualNetworkName = "${NameStem}_Infra_net"
$SecurityGroupName  = "${NameStem}_Infra_nsg"
$StorageType        = $EnvConfig.Azure.StorageType
$ServerPublisherName = $EnvConfig.Azure.ServerImage.Publisher
$ServerOffer         = $EnvConfig.Azure.ServerImage.Offer
$ServerSkus          = $EnvConfig.Azure.ServerImage.Sku
$ClientPublisherName = $EnvConfig.Azure.ClientImage.Publisher
$ClientOffer         = $EnvConfig.Azure.ClientImage.Offer
$ClientSkus          = $EnvConfig.Azure.ClientImage.Sku
$MachineNames = @("$NameStem-DC1", "$NameStem-DC2", "$NameStem-RTR", "$NameStem-EXCH", "$NameStem-1")
$VirtualMachineCredential = $null
$MachineSKU         = $EnvConfig.Azure.VmSizes.Default
$ExchangeMachineSKU = $EnvConfig.Azure.VmSizes.Exchange
$ClientMachineSKU   = $EnvConfig.Azure.VmSizes.Client
$DataDiskSize       = $EnvConfig.Azure.DataDiskGB
$RgTags = $EnvConfig.Azure.BaseTags + @{ Department = $NameStem; Owner = $Owner }
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptTitle = "$NameStem Azure Resource Creation Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $NameStem + "_new_Azure_resource_creation_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-LogFile -LogFile $LogFile -LogString "Processing commenced, running as user '$NameStem\$env:USERNAME'"
Write-LogFile -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-LogFile -LogFile $LogFile -LogString "$ScriptTitle"
Write-LogFile -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "

function Remove-IncompleteBuildResource {
    # Opt-in reconciliation for abandoned or permanently-failing builds. For each
    # expected machine with NO completed VM, remove the standalone resources its
    # build would have left behind so they don't bill indefinitely. Best-effort,
    # dependency-ordered (NIC releases the public IP, then disks); never touches a
    # machine whose VM exists. Naming mirrors the build loop exactly.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$MachineNames
        ,[Parameter(Mandatory)][string]$ResourceGroupName
        ,[Parameter(Mandatory)][string]$LogFile
    )
    foreach ($Machine in $MachineNames) {
        if (Get-AzVM -ResourceGroupName $ResourceGroupName -Name $Machine -ErrorAction SilentlyContinue) {
            Write-LogFile -LogFile $LogFile -LogString "Cleanup: $Machine has a VM - leaving its resources intact"
            continue
        }
        Write-LogFile -LogFile $LogFile -LogString "Cleanup: $Machine has no VM - removing any orphaned build resources" -ForegroundColor Yellow
        $nicName = $Machine + "_NIC"
        $pipName = $Machine + "_PIP"
        try {
            if (Get-AzNetworkInterface -Name $nicName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue) {
                Remove-AzNetworkInterface -Name $nicName -ResourceGroupName $ResourceGroupName -Force -ErrorAction Stop
                Write-LogFile -LogFile $LogFile -LogString "Cleanup: removed NIC $nicName"
            }
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "Cleanup WARNING: NIC $nicName : $_" -ForegroundColor Yellow
        }
        try {
            if (Get-AzPublicIpAddress -Name $pipName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue) {
                Remove-AzPublicIpAddress -Name $pipName -ResourceGroupName $ResourceGroupName -Force -ErrorAction Stop
                Write-LogFile -LogFile $LogFile -LogString "Cleanup: removed Public IP $pipName"
            }
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "Cleanup WARNING: Public IP $pipName : $_" -ForegroundColor Yellow
        }
        foreach ($disk in @(($Machine + "_DataDisk_0"), ($Machine + "_OSDisk"))) {
            try {
                if (Get-AzDisk -DiskName $disk -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue) {
                    Remove-AzDisk -DiskName $disk -ResourceGroupName $ResourceGroupName -Force -ErrorAction Stop | Out-Null
                    Write-LogFile -LogFile $LogFile -LogString "Cleanup: removed disk $disk"
                }
            } catch {
                Write-LogFile -LogFile $LogFile -LogString "Cleanup WARNING: disk $disk : $_" -ForegroundColor Yellow
            }
        }
    }
}

try {
    try {
        Connect-AzAccount -SubscriptionId $SubscriptionId -TenantID $TenantID
    } catch {
        Write-LogFile -LogFile $LogFile -LogString "ERROR Connecting to Azure : $_" -ForegroundColor Red
        throw
    }
    try {
        Set-AzContext -SubscriptionId $SubscriptionId -TenantID $TenantID
    } catch {
        Write-LogFile -LogFile $LogFile -LogString "ERROR Setting Az context : $_" -ForegroundColor Red
        throw
    }
    Write-LogFile -LogFile $LogFile -LogString "Verifying the Azure login subscription status..."
    if(-not $(Get-AzContext)) {
        Write-LogFile -LogFile $LogFile -LogString "Login to Azure subscription failed, no valid subscription found."
        try {
            Connect-AzAccount -SubscriptionId $SubscriptionId -TenantID $TenantID
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "ERROR Connecting to Azure : $_" -ForegroundColor Red
            throw
        }
        try{
            Set-AzContext -SubscriptionId $SubscriptionId -TenantID $TenantID
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "ERROR Setting Az context : $_" -ForegroundColor Red
            throw
        }
    }
    if ($RemoveIncomplete) {
        Write-LogFile -LogFile $LogFile -LogString "RemoveIncomplete mode: reconciling orphaned build resources"
        Remove-IncompleteBuildResource -MachineNames $MachineNames -ResourceGroupName $ResourceGroupName -LogFile $LogFile
        Write-LogFile -LogFile $LogFile -LogString "RemoveIncomplete mode: complete"
        return
    }
    try {
        $ResourceGroup = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
    } catch {
        Write-LogFile -LogFile $LogFile -LogString "ERROR getting $ResourceGroupName : $_" -ForegroundColor Red
    }
    if (!($ResourceGroup)) {
        Write-LogFile -LogFile $LogFile -LogString "Creating new resource group $ResourceGroupName"
        try {
            $ResourceGroup = New-AzResourceGroup -Name $ResourceGroupName -Location $Location -Tag $RgTags
            Write-LogFile -LogFile $LogFile -LogString "New resource group $ResourceGroupName created"
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "ERROR creating $ResourceGroupName : $_" -ForegroundColor Red
            throw
        }
    } else {
        Write-LogFile -LogFile $LogFile -LogString "$ResourceGroupName Resource Group already exists, skipping new resource group creation..."
    }
    try {
        $VirtualNetwork = Get-AzVirtualNetwork -Name $VirtualNetworkName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
    } catch {
        Write-LogFile -LogFile $LogFile -LogString "ERROR getting $VirtualNetworkName : $_" -ForegroundColor Red
    }
    if (!($VirtualNetwork)) {
        Write-LogFile -LogFile $LogFile -LogString "Creating new virtual network $VirtualNetworkName"
        try {
            $SecurityGroup = Get-AzNetworkSecurityGroup -Name $SecurityGroupName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "ERROR getting $SecurityGroupName : $_" -ForegroundColor Red
        }
        if (!($SecurityGroup)) {
            try {
                $nsgRuleRDP = New-AzNetworkSecurityRuleConfig -Name rdp-rule -Description "Allow RDP" -Access Allow -Protocol Tcp -Direction Inbound -Priority 100 -SourceAddressPrefix $MyIPAddress -SourcePortRange * -DestinationAddressPrefix * -DestinationPortRange 3389
                Write-LogFile -LogFile $LogFile -LogString "New RDP Security rule created"
            } catch {
                Write-LogFile -LogFile $LogFile -LogString "ERROR creating RDP Security rule : $_" -ForegroundColor Red
                throw
            }
            try {
                $SecurityGroup = New-AzNetworkSecurityGroup -Name $SecurityGroupName -ResourceGroupName $ResourceGroupName -Location $Location -SecurityRules $nsgRuleRDP
                Write-LogFile -LogFile $LogFile -LogString "New security group $SecurityGroupName created"
            } catch {
                Write-LogFile -LogFile $LogFile -LogString "ERROR creating $SecurityGroupName : $_" -ForegroundColor Red
                throw
            }
        }
        For ($i=0; $i -le $IPNet.Length -1; $i++) {
            $ScopeNetwork = $IPNet[$i] + ".0"
            $ScopeName = "$ScopeNetwork/$SubnetCDIR"
            try {
                $Subnet += New-AzVirtualNetworkSubnetConfig -Name $Site[$i] -AddressPrefix $ScopeName -NetworkSecurityGroup $SecurityGroup
                Write-LogFile -LogFile $LogFile -LogString "New subnet $ScopeName created"
            } catch {
                Write-LogFile -LogFile $LogFile -LogString "ERROR creating Subnet : $_" -ForegroundColor Red
                throw
            }
        }
        try {
            $VirtualNetwork = New-AzVirtualNetwork -Name $VirtualNetworkName -ResourceGroupName $ResourceGroupName -Location $Location -AddressPrefix $VnetPrefix -Subnet ($Subnet) -DnsServer $VnetDNSServer
            Write-LogFile -LogFile $LogFile -LogString "New virtual network $VirtualNetworkName created"
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "ERROR creating $VirtualNetworkName : $_" -ForegroundColor Red
            throw
        }
    } else {
        Write-LogFile -LogFile $LogFile -LogString "$VirtualNetworkName virtual network already exists, skipping new virtual network creation..."
    }
    $VirtualMachineCredential = Get-Credential -Message 'Please enter the vm credentials'
    foreach ($Machine in $MachineNames) {
        $VirtualMachineName = $Machine
        $VM = $null
        $VirtualMachineNIC = $null
        $DataDisk = $null
        $PublicIPAddress = $null
        try {
            $VM = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $VirtualMachineName -ErrorAction SilentlyContinue
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "ERROR getting $VirtualMachineName : $_" -ForegroundColor Red
        }
        if (!($VM)) {
            try {
                $OSDiskName = $VirtualMachineName + "_OSDisk"
                $DataDiskName = $VirtualMachineName + "_DataDisk_0"
                $VirtualMachineNICName = $VirtualMachineName + "_NIC"
                $VirtualMachineSize = $MachineSKU
                try {
                    $VirtualMachineNIC = Get-AzNetworkInterface -Name $VirtualMachineNICName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
                } catch {
                    Write-LogFile -LogFile $LogFile -LogString "ERROR getting $VirtualMachineNICName : $_" -ForegroundColor Red
                }
                if (!($VirtualMachineNIC)) {
                    try {
                        $VirtualMachineNIC = New-AzNetworkInterface -Name $VirtualMachineNICName -ResourceGroupName $ResourceGroupName -Location $Location -SubnetId $VirtualNetwork.Subnets[0].Id
                        Write-LogFile -LogFile $LogFile -LogString "New VM NIC $VirtualMachineNICName created"
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "ERROR creating $VirtualMachineNICName : $_" -ForegroundColor Red
                        throw
                    }
                }
                if ($VirtualMachineName -eq "$NameStem-EXCH") {
                    $VirtualMachineSize = $ExchangeMachineSKU
                    $PublicIPAddressName = $VirtualMachineName + "_PIP"
                    try {
                        $PublicIPAddress = Get-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "ERROR getting $PublicIPAddressName : $_" -ForegroundColor Red
                    }
                    if (!($PublicIPAddress)) {
                        try {
                            $PublicIPAddress = New-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName -Location $Location -AllocationMethod Static
                            Write-LogFile -LogFile $LogFile -LogString "New VM Public IP $PublicIPAddressName created"
                        } catch {
                            Write-LogFile -LogFile $LogFile -LogString "ERROR creating $PublicIPAddressName : $_" -ForegroundColor Red
                            throw
                        }
                    }
                    try {
                        $VirtualMachineNIC | Set-AzNetworkInterfaceIpConfig -Name $VirtualMachineNIC.IpConfigurations[0].Name -PublicIpAddressId $PublicIPAddress.Id | Set-AzNetworkInterface
                        Write-LogFile -LogFile $LogFile -LogString "Adding $PublicIPAddressName to $VirtualMachineNICName"
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "ERROR setting $VirtualMachineNICName IP config : $_" -ForegroundColor Red
                        throw
                    }
                    try {
                        $VirtualMachine_PublicIP = (Get-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName).IpAddress
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "ERROR getting $PublicIPAddressName : $_" -ForegroundColor Red
                    }
                    try {
                        $VirtualMachine_PrivateIP = (Get-AzNetworkInterface -Name $VirtualMachineNICName -ResourceGroupName $ResourceGroupName).IpConfigurations.PrivateIpAddress
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "ERROR getting VM private IP: $_" -ForegroundColor Red
                    }
                    try {
                        $SecurityGroup = Get-AzNetworkSecurityGroup -Name $SecurityGroupName -ResourceGroupName $ResourceGroupName
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "ERROR getting $SecurityGroupName : $_" -ForegroundColor Red
                    }
                    if (!(Get-AzNetworkSecurityRuleConfig -Name http-rule -NetworkSecurityGroup $SecurityGroup -ErrorAction SilentlyContinue)) {
                        try {
                            $SecurityGroup | Add-AzNetworkSecurityRuleConfig -Name http-rule -Description "Allow HTTP" -Access Allow -Protocol Tcp -Direction Inbound -Priority 110 -SourceAddressPrefix * -SourcePortRange * -DestinationAddressPrefix $VirtualMachine_PrivateIP -DestinationPortRange 80
                            Write-LogFile -LogFile $LogFile -LogString "Creating HTTP rule for $SecurityGroupName"
                        } catch {
                            Write-LogFile -LogFile $LogFile -LogString "ERROR setting $SecurityGroupName rule : $_" -ForegroundColor Red
                            throw
                        }
                        try {
                            $SecurityGroup | Set-AzNetworkSecurityGroup
                            Write-LogFile -LogFile $LogFile -LogString "Adding HTTP rule to $SecurityGroupName"
                        } catch {
                            Write-LogFile -LogFile $LogFile -LogString "ERROR setting $SecurityGroupName : $_" -ForegroundColor Red
                            throw
                        }
                    }
                    if (!(Get-AzNetworkSecurityRuleConfig -Name https-rule -NetworkSecurityGroup $SecurityGroup -ErrorAction SilentlyContinue)) {
                        try {
                            $SecurityGroup | Add-AzNetworkSecurityRuleConfig -Name https-rule -Description "Allow HTTPS" -Access Allow -Protocol Tcp -Direction Inbound -Priority 120 -SourceAddressPrefix * -SourcePortRange * -DestinationAddressPrefix $VirtualMachine_PrivateIP -DestinationPortRange 443
                            Write-LogFile -LogFile $LogFile -LogString "Creating HTTPS rule for $SecurityGroupName"
                        } catch {
                            Write-LogFile -LogFile $LogFile -LogString "ERROR setting $SecurityGroupName rule : $_" -ForegroundColor Red
                            throw
                        }
                        try {
                            $SecurityGroup | Set-AzNetworkSecurityGroup
                            Write-LogFile -LogFile $LogFile -LogString "Adding HTTPS rule to $SecurityGroupName"
                        } catch {
                            Write-LogFile -LogFile $LogFile -LogString "ERROR setting $SecurityGroupName : $_" -ForegroundColor Red
                            throw
                        }
                    }
                }
                if ($VirtualMachineName -eq "$NameStem-1") {
                    $VirtualMachineSize = $ClientMachineSKU
                    $PublicIPAddressName = $VirtualMachineName + "_PIP"
                    try {
                        $PublicIPAddress = Get-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "ERROR getting $PublicIPAddressName : $_" -ForegroundColor Red
                    }
                    if (!($PublicIPAddress)) {
                        try {
                            $PublicIPAddress = New-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName -Location $Location -AllocationMethod Static
                            Write-LogFile -LogFile $LogFile -LogString "New VM Public IP $PublicIPAddressName created"
                        } catch {
                            Write-LogFile -LogFile $LogFile -LogString "ERROR creating $PublicIPAddressName : $_" -ForegroundColor Red
                            throw
                        }
                    }
                    try {
                        $VirtualMachineNIC | Set-AzNetworkInterfaceIpConfig -Name $VirtualMachineNIC.IpConfigurations[0].Name -PublicIpAddressId $PublicIPAddress.Id | Set-AzNetworkInterface
                        Write-LogFile -LogFile $LogFile -LogString "Adding $PublicIPAddressName to $VirtualMachineNICName"
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "ERROR setting $VirtualMachineNICName config : $_" -ForegroundColor Red
                        throw
                    }
                }
                try {
                    $VirtualMachine = New-AzVMConfig -VMName $VirtualMachineName -VMSize $virtualMachineSize -Tags $RgTags
                    Write-LogFile -LogFile $LogFile -LogString "Setting $VirtualMachineName tags"
                } catch {
                    Write-LogFile -LogFile $LogFile -LogString "ERROR setting $VirtualMachineName tags : $_" -ForegroundColor Red
                    throw
                }
                try {
                    $VirtualMachine = Set-AzVMOperatingSystem -VM $VirtualMachine -Windows -ComputerName $VirtualMachineName -Credential $VirtualMachineCredential -ProvisionVMAgent -EnableAutoUpdate
                    Write-LogFile -LogFile $LogFile -LogString "Setting $VirtualMachineName OS config"
                } catch {
                    Write-LogFile -LogFile $LogFile -LogString "ERROR setting $VirtualMachineName OS config : $_" -ForegroundColor Red
                    throw
                }
                try {
                    $VirtualMachine = Add-AzVMNetworkInterface -VM $VirtualMachine -Id $VirtualMachineNIC.Id
                    Write-LogFile -LogFile $LogFile -LogString "Setting $VirtualMachineNICName"
                } catch {
                    Write-LogFile -LogFile $LogFile -LogString "ERROR setting $VirtualMachineNICName : $_" -ForegroundColor Red
                    throw
                }
                try {
                    $VirtualMachine = Set-AzVMBootDiagnostic -VM $VirtualMachine -Enable -ResourceGroupName $ResourceGroupName
                    Write-LogFile -LogFile $LogFile -LogString "Setting $VirtualMachineName boot diagnostics"
                } catch {
                    Write-LogFile -LogFile $LogFile -LogString "ERROR setting $VirtualMachineName boot diagnostics : $_" -ForegroundColor Red
                    throw
                }
                if ($VirtualMachineName -eq "$NameStem-DC1" -or $VirtualMachineName -eq "$NameStem-DC2" -or $VirtualMachineName -eq "$NameStem-RTR" -or $VirtualMachineName -eq "$NameStem-EXCH") {
                    try {
                        $VirtualMachine = Set-AzVMSourceImage -VM $VirtualMachine -PublisherName $ServerPublisherName -Offer $ServerOffer -Skus $ServerSkus -Version "latest"
                        Write-LogFile -LogFile $LogFile -LogString "Setting $VirtualMachineName source image"
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "ERROR setting $VirtualMachineName source image : $_" -ForegroundColor Red
                        throw
                    }
                    if ($VirtualMachineName -eq "$NameStem-DC1" -or $VirtualMachineName -eq "$NameStem-DC2") {
                        try {
                            $DataDisk = Get-AzDisk -Name $DataDiskName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
                        } catch {
                            Write-LogFile -LogFile $LogFile -LogString "ERROR getting $DataDiskName : $_" -ForegroundColor Red
                        }
                        if (!($DataDisk)) {
                            try {
                                $DataDiskConfig = New-AzDiskConfig -SkuName $StorageType -Location $Location -CreateOption Empty -DiskSizeGB $DataDiskSize
                                Write-LogFile -LogFile $LogFile -LogString "Setting $DataDiskName config"
                            } catch {
                                Write-LogFile -LogFile $LogFile -LogString "ERROR setting $DataDiskName config : $_" -ForegroundColor Red
                                throw
                            }
                            try {
                                $DataDisk = New-AzDisk -DiskName $DataDiskName -Disk $DataDiskConfig -ResourceGroupName $ResourceGroupName
                                Write-LogFile -LogFile $LogFile -LogString "Creating $DataDiskName"
                            } catch {
                                Write-LogFile -LogFile $LogFile -LogString "ERROR creating $DataDiskName : $_" -ForegroundColor Red
                                throw
                            }
                        }
                        try {
                            $VirtualMachine = Add-AzVMDataDisk -VM $VirtualMachine -Name $DataDiskName -CreateOption Attach -ManagedDiskId $DataDisk.Id -Lun 0
                            Write-LogFile -LogFile $LogFile -LogString "Adding $DataDiskName to $VirtualMachineName"
                        } catch {
                            Write-LogFile -LogFile $LogFile -LogString "ERROR adding $DataDiskName to $VirtualMachineName : $_" -ForegroundColor Red
                            throw
                        }
                    }
                } else {
                    try {
                        $VirtualMachine = Set-AzVMSourceImage -VM $VirtualMachine -PublisherName $ClientPublisherName -Offer $ClientOffer -Skus $ClientSkus -Version "latest"
                        Write-LogFile -LogFile $LogFile -LogString "Setting $VirtualMachineName source image"
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "ERROR setting $VirtualMachineName source image : $_" -ForegroundColor Red
                        throw
                    }
                }
                try {
                    $VirtualMachine = Set-AzVMOSDisk -VM $VirtualMachine -Name $OSDiskName -StorageAccountType $StorageType -Caching ReadWrite -CreateOption FromImage
                    Write-LogFile -LogFile $LogFile -LogString "Setting $VirtualMachineName OS disk"
                } catch {
                    Write-LogFile -LogFile $LogFile -LogString "ERROR setting $VirtualMachineName OS disk : $_" -ForegroundColor Red
                    throw
                }
                try {
                    New-AzVM -VM $VirtualMachine -ResourceGroupName $ResourceGroupName -Location $location
                    Write-LogFile -LogFile $LogFile -LogString "Creating $VirtualMachineName"
                } catch {
                    Write-LogFile -LogFile $LogFile -LogString "ERROR creating $VirtualMachineName : $_" -ForegroundColor Red
                    throw
                }
            } catch {
                Write-LogFile -LogFile $LogFile -LogString "ERROR creating $VirtualMachineName : $_" -ForegroundColor Red
                continue
            }
        } else {
            Write-LogFile -LogFile $LogFile -LogString "VM $VirtualMachineName already exists, skipping new VM creation..."
        }
    }
} finally {
    try {
        # Clear sensitive variables securely
        Write-LogFile -LogFile $LogFile -LogString "Clearing credentials and disconnecting session"
        if ($VirtualMachineCredential) {
            $VirtualMachineCredential.Password.Dispose()
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    } finally {
        Disconnect-AzAccount -ErrorAction SilentlyContinue -confirm:$false | Out-Null
        Clear-AzContext -ErrorAction SilentlyContinue -force | Out-Null
    }
}
