#Requires -Modules Az

[CmdletBinding()]
param(
    [string]$SubscriptionId
    , [string]$TenantID
    , [string]$MyIPAddress
    , [string]$NameStem
    , [string]$Owner
    , [string]$LogFile
)

Set-StrictMode -Version Latest

. $PSScriptRoot\helpers.ps1

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
$IPNet = @("10.71.104", "10.71.105")
$Site = @("Site1", "Site2")
$Subnet = @()
$SubnetCDIR = "24"
$VnetCDIR = "21"
$DNSServer4 = "4"
$VnetDNSServer = $IPNet[0] + "." + $DNSServer4
$VnetPrefix = $IPNet[0] + ".0/" + $VnetCDIR
$NameStemLower = $NameStem.ToLower()
$ResourceGroupName = $NameStem + "_Infra_RG"
$Location = "UK South"
$VirtualNetworkName = $NameStem + "_Infra_net"
$SecurityGroupName = $NameStem + "_Infra_nsg"
$StorageType = "StandardSSD_LRS"
$ServerPublisherName = "MicrosoftWindowsServer"
$ServerOffer = "WindowsServer"
$ServerSkus = "2025-Datacenter-g2"
$ClientPublisherName = "MicrosoftWindowsDesktop"
$ClientOffer = "windows-11"
$ClientSkus = "win11-25h2-pro"
$MachineNames = @("$NameStem-DC1", "$NameStem-DC2", "$NameStem-RTR", "$NameStem-EXCH", "$NameStem-1")
$VirtualMachineCredential = Get-Credential -Message 'Please enter the vm credentials'
$MachineSKU = "Standard_B2s"
$ExchangeMachineSKU = "Standard_D4s_v5"
$ClientMachineSKU = "Standard_D2s_v5"
$DataDiskSize = 64
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptTitle = "$NameStem Azure Resource Creation Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
if (!$LogFile) {
    $LogFileName = $NameStem + "_new_Azure_resource_creation_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
    Write-Log -LogFile $LogFile -LogString ("=" * 80)
    Write-Log -LogFile $LogFile -LogString "Log file is '$LogFile'"
    Write-Log -LogFile $LogFile -LogString "Processing commenced, running as user '$NameStem\$env:USERNAME'"
    Write-Log -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
    Write-Log -LogFile $LogFile -LogString "$ScriptTitle"
    Write-Log -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
    Write-Log -LogFile $LogFile -LogString ("=" * 80)
    Write-Log -LogFile $LogFile -LogString " "
}
try {
    Connect-AzAccount -SubscriptionId $SubscriptionId -TenantID $TenantID
} catch {
    Write-Log -LogFile $LogFile -LogString "ERROR Connecting to Azure : $_" -ForegroundColor Red
    Exit
}
try {
    Set-AzContext -SubscriptionId $SubscriptionId -TenantID $TenantID
} catch {
    Write-Log -LogFile $LogFile -LogString "ERROR Setting Az context : $_" -ForegroundColor Red
    Exit
}
Write-Log -LogFile $LogFile -LogString "Verifying the Azure login subscription status..."
if(-not $(Get-AzContext)) {
    Write-Log -LogFile $LogFile -LogString "Login to Azure subscription failed, no valid subscription found."
    try {
        Connect-AzAccount -SubscriptionId $SubscriptionId -TenantID $TenantID
    } catch {
        Write-Log -LogFile $LogFile -LogString "ERROR Connecting to Azure : $_" -ForegroundColor Red
        Exit
    }
    try{
        Set-AzContext -SubscriptionId $SubscriptionId -TenantID $TenantID
    } catch {
        Write-Log -LogFile $LogFile -LogString "ERROR Setting Az context : $_" -ForegroundColor Red
        Exit
    }
}

try {
    $ResourceGroup = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
} catch {
    Write-Log -LogFile $LogFile -LogString "ERROR getting $ResourceGroupName : $_" -ForegroundColor Red
}
if (!($ResourceGroup)) {
    Write-Log -LogFile $LogFile -LogString "Creating new resource group $ResourceGroupName"
    try {
        $ResourceGroup = New-AzResourceGroup -Name $ResourceGroupName -Location $Location -Tag @{Criticality = "Tier 1"; Department = "$NameStem"; Environment = "Production"; Owner = $Owner}
        Write-Log -LogFile $LogFile -LogString "New resource group $ResourceGroupName created"
    } catch {
        Write-Log -LogFile $LogFile -LogString "ERROR creating $ResourceGroupName : $_" -ForegroundColor Red
        Exit
    }
} else {
    Write-Log -LogFile $LogFile -LogString "$ResourceGroupName Resource Group already exists, skipping new resource group creation..."
}

try {
    $VirtualNetwork = Get-AzVirtualNetwork -Name $VirtualNetworkName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
} catch {
    Write-Log -LogFile $LogFile -LogString "ERROR getting $VirtualNetworkName : $_" -ForegroundColor Red
}
if (!($VirtualNetwork)) {
    Write-Log -LogFile $LogFile -LogString "Creating new virtual network $VirtualNetworkName"
    try {
        $SecurityGroup = Get-AzNetworkSecurityGroup -Name $SecurityGroupName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
    } catch {
        Write-Log -LogFile $LogFile -LogString "ERROR getting $SecurityGroupName : $_" -ForegroundColor Red
    }
    if (!($SecurityGroup)) {
        try {
            $nsgRuleRDP = New-AzNetworkSecurityRuleConfig -Name rdp-rule -Description "Allow RDP" -Access Allow -Protocol Tcp -Direction Inbound -Priority 100 -SourceAddressPrefix $MyIPAddress -SourcePortRange * -DestinationAddressPrefix * -DestinationPortRange 3389
            Write-Log -LogFile $LogFile -LogString "New RDP Security rule created"
        } catch {
            Write-Log -LogFile $LogFile -LogString "ERROR creating RDP Security rule : $_" -ForegroundColor Red
            Exit
        }
        try {
            $SecurityGroup = New-AzNetworkSecurityGroup -Name $SecurityGroupName -ResourceGroupName $ResourceGroupName -Location $Location -SecurityRules $nsgRuleRDP
            Write-Log -LogFile $LogFile -LogString "New security group $SecurityGroupName created"
        } catch {
            Write-Log -LogFile $LogFile -LogString "ERROR creating $SecurityGroupName : $_" -ForegroundColor Red
            Exit
        }
    }
    For ($i=0; $i -le $IPNet.Length -1; $i++) {
        $ScopeNetwork = $IPNet[$i] + ".0"
        $ScopeName = "$ScopeNetwork/$SubnetCDIR"
        try {
            $Subnet += New-AzVirtualNetworkSubnetConfig -Name $Site[$i] -AddressPrefix $ScopeName -NetworkSecurityGroup $SecurityGroup
            Write-Log -LogFile $LogFile -LogString "New subnet $ScopeName created"
        } catch {
            Write-Log -LogFile $LogFile -LogString "ERROR creating Subnet : $_" -ForegroundColor Red
            Exit
        }
    }
    try {
        $VirtualNetwork = New-AzVirtualNetwork -Name $VirtualNetworkName -ResourceGroupName $ResourceGroupName -Location $Location -AddressPrefix $VnetPrefix -Subnet ($Subnet) -DnsServer $VnetDNSServer
        Write-Log -LogFile $LogFile -LogString "New virtual network $VirtualNetworkName created"
    } catch {
        Write-Log -LogFile $LogFile -LogString "ERROR creating $VirtualNetworkName : $_" -ForegroundColor Red
        Exit
    }
} else {
    Write-Log -LogFile $LogFile -LogString "$VirtualNetworkName virtual network already exists, skipping new virtual network creation..."
}

foreach ($Machine in $MachineNames) {
    $VirtualMachineName = $Machine
    $VM = $null
    $VirtualMachineNIC = $null
    $DataDisk = $null
    $PublicIPAddress = $null
    try {
        $VM = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $VirtualMachineName -ErrorAction SilentlyContinue
    } catch {
        Write-Log -LogFile $LogFile -LogString "ERROR getting $VirtualMachineName : $_" -ForegroundColor Red
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
                Write-Log -LogFile $LogFile -LogString "ERROR getting $VirtualMachineNICName : $_" -ForegroundColor Red
            }
            if (!($VirtualMachineNIC)) {
                try {
                    $VirtualMachineNIC = New-AzNetworkInterface -Name $VirtualMachineNICName -ResourceGroupName $ResourceGroupName -Location $Location -SubnetId $VirtualNetwork.Subnets[0].Id
                    Write-Log -LogFile $LogFile -LogString "New VM NIC $VirtualMachineNICName created"
                } catch {
                    Write-Log -LogFile $LogFile -LogString "ERROR creating $VirtualMachineNICName : $_" -ForegroundColor Red
                }
            }
            if ($VirtualMachineName -eq "$NameStem-EXCH") {
                $VirtualMachineSize = $ExchangeMachineSKU
                $PublicIPAddressName = $VirtualMachineName + "_PIP"
                try {
                    $PublicIPAddress = Get-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
                } catch {
                    Write-Log -LogFile $LogFile -LogString "ERROR getting $PublicIPAddressName : $_" -ForegroundColor Red
                }
                if (!($PublicIPAddress)) {
                    try {
                        $PublicIPAddress = New-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName -Location $Location -AllocationMethod Static
                        Write-Log -LogFile $LogFile -LogString "New VM Public IP $PublicIPAddressName created"
                    } catch {
                        Write-Log -LogFile $LogFile -LogString "ERROR creating $PublicIPAddressName : $_" -ForegroundColor Red
                    }
                }
                try {
                    $VirtualMachineNIC | Set-AzNetworkInterfaceIpConfig -Name $VirtualMachineNIC.IpConfigurations[0].Name -PublicIpAddressId $PublicIPAddress.Id | Set-AzNetworkInterface
                    Write-Log -LogFile $LogFile -LogString "Adding $PublicIPAddressName to $VirtualMachineNICName created"
                } catch {
                    Write-Log -LogFile $LogFile -LogString "ERROR setting $VirtualMachineNICName IP config : $_" -ForegroundColor Red
                }
                try {
                    $VirtualMachine_PublicIP = (Get-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName).IpAddress
                } catch {
                    Write-Log -LogFile $LogFile -LogString "ERROR getting $PublicIPAddressName : $_" -ForegroundColor Red
                }
                try {
                    $VirtualMachine_PrivateIP = (Get-AzNetworkInterface -Name $VirtualMachineNICName -ResourceGroupName $ResourceGroupName).IpConfigurations.PrivateIpAddress
                } catch {
                    Write-Log -LogFile $LogFile -LogString "ERROR getting VM private IP: $_" -ForegroundColor Red
                }
                $VirtualMachineNIC_IP = @("$VirtualMachine_PrivateIP","$VirtualMachine_PublicIP")
                try {
                    $SecurityGroup = Get-AzNetworkSecurityGroup -Name $SecurityGroupName -ResourceGroupName $ResourceGroupName
                } catch {
                    Write-Log -LogFile $LogFile -LogString "ERROR getting $SecurityGroupName : $_" -ForegroundColor Red
                }
                if (!(Get-AzNetworkSecurityRuleConfig -Name http-rule -NetworkSecurityGroup $SecurityGroup)) {
                    try {
                        $SecurityGroup | Add-AzNetworkSecurityRuleConfig -Name http-rule -Description "Allow HTTP" -Access Allow -Protocol Tcp -Direction Inbound -Priority 110 -SourceAddressPrefix * -SourcePortRange * -DestinationAddressPrefix $VirtualMachineNIC_IP -DestinationPortRange 80
                        Write-Log -LogFile $LogFile -LogString "Creating HTTP rule to $SecurityGroupName"
                    } catch {
                        Write-Log -LogFile $LogFile -LogString "ERROR setting $SecurityGroupName rule : $_" -ForegroundColor Red
                    }
                    try {
                        $SecurityGroup | Set-AzNetworkSecurityGroup
                        Write-Log -LogFile $LogFile -LogString "Adding HTTP rule to $SecurityGroupName"
                    } catch {
                        Write-Log -LogFile $LogFile -LogString "ERROR setting $SecurityGroupName : $_" -ForegroundColor Red
                    }
                }
                if (!(Get-AzNetworkSecurityRuleConfig -Name https-rule -NetworkSecurityGroup $SecurityGroup)) {
                    try {
                        $SecurityGroup | Add-AzNetworkSecurityRuleConfig -Name https-rule -Description "Allow HTTPS" -Access Allow -Protocol Tcp -Direction Inbound -Priority 120 -SourceAddressPrefix * -SourcePortRange * -DestinationAddressPrefix $VirtualMachineNIC_IP -DestinationPortRange 443
                        Write-Log -LogFile $LogFile -LogString "Creating HTTP rule to $SecurityGroupName"
                    } catch {
                        Write-Log -LogFile $LogFile -LogString "ERROR setting $SecurityGroupName rule : $_" -ForegroundColor Red
                    }
                    try {
                        $SecurityGroup | Set-AzNetworkSecurityGroup
                        Write-Log -LogFile $LogFile -LogString "Adding HTTPS rule to $SecurityGroupName"
                    } catch {
                        Write-Log -LogFile $LogFile -LogString "ERROR setting $SecurityGroupName : $_" -ForegroundColor Red
                    }
                }
            }
            if ($VirtualMachineName -eq "$NameStem-1") {
                $VirtualMachineSize = $ClientMachineSKU
                $PublicIPAddressName = $VirtualMachineName + "_PIP"
                try {
                    $PublicIPAddress = Get-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
                } catch {
                    Write-Log -LogFile $LogFile -LogString "ERROR getting $PublicIPAddressName : $_" -ForegroundColor Red
                }
                if (!($PublicIPAddress)) {
                    try {
                        $PublicIPAddress = New-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName -Location $Location -AllocationMethod Static
                        Write-Log -LogFile $LogFile -LogString "New VM Public IP $PublicIPAddressName created"
                    } catch {
                        Write-Log -LogFile $LogFile -LogString "ERROR creating $PublicIPAddressName : $_" -ForegroundColor Red
                    }
                }
                try {
                    $VirtualMachineNIC | Set-AzNetworkInterfaceIpConfig -Name $VirtualMachineNIC.IpConfigurations[0].Name -PublicIpAddressId $PublicIPAddress.Id | Set-AzNetworkInterface
                    Write-Log -LogFile $LogFile -LogString "Adding $PublicIPAddressName to $VirtualMachineNICName created"
                } catch {
                    Write-Log -LogFile $LogFile -LogString "ERROR setting $VirtualMachineNICName config : $_" -ForegroundColor Red
                }
            }
            try {
                $VirtualMachine = New-AzVMConfig -VMName $VirtualMachineName -VMSize $virtualMachineSize -Tags @{Criticality = "Tier 1"; Environment = "Production"}
                Write-Log -LogFile $LogFile -LogString "Setting $VirtualMachineName tags"
            } catch {
                Write-Log -LogFile $LogFile -LogString "ERROR setting $VirtualMachineName tags : $_" -ForegroundColor Red
            }
            try {
                $VirtualMachine = Set-AzVMOperatingSystem -VM $VirtualMachine -Windows -ComputerName $VirtualMachineName -Credential $VirtualMachineCredential -ProvisionVMAgent -EnableAutoUpdate
                Write-Log -LogFile $LogFile -LogString "Setting $VirtualMachineName OS config"
            } catch {
                Write-Log -LogFile $LogFile -LogString "ERROR setting $VirtualMachineName OS config : $_" -ForegroundColor Red
            }
            try {
                $VirtualMachine = Add-AzVMNetworkInterface -VM $VirtualMachine -Id $VirtualMachineNIC.Id
                Write-Log -LogFile $LogFile -LogString "Setting $VirtualMachineNICName"
            } catch {
                Write-Log -LogFile $LogFile -LogString "ERROR setting $VirtualMachineNICName : $_" -ForegroundColor Red
            }
            try {
                $VirtualMachine = Set-AzVMBootDiagnostic -VM $VirtualMachine -Enable -ResourceGroupName $ResourceGroupName
                Write-Log -LogFile $LogFile -LogString "Setting $VirtualMachineName boot diagnostics"
            } catch {
                Write-Log -LogFile $LogFile -LogString "ERROR setting $VirtualMachineName boot diagnostics : $_" -ForegroundColor Red
            }
            if ($VirtualMachineName -eq "$NameStem-DC1" -or $VirtualMachineName -eq "$NameStem-DC2" -or $VirtualMachineName -eq "$NameStem-RTR" -or $VirtualMachineName -eq "$NameStem-EXCH") {
                try {
                    $VirtualMachine = Set-AzVMSourceImage -VM $VirtualMachine -PublisherName $ServerPublisherName -Offer $ServerOffer -Skus $ServerSkus -Version "latest"
                    Write-Log -LogFile $LogFile -LogString "Setting $VirtualMachineName source image"
                } catch {
                    Write-Log -LogFile $LogFile -LogString "ERROR setting $VirtualMachineName source image : $_" -ForegroundColor Red
                }
                if ($VirtualMachineName -eq "$NameStem-DC1" -or $VirtualMachineName -eq "$NameStem-DC2") {
                    try {
                        $DataDisk = Get-AzDisk -Name $DataDiskName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
                    } catch {
                        Write-Log -LogFile $LogFile -LogString "ERROR getting $DataDiskName : $_" -ForegroundColor Red
                    }
                    if (!($DataDisk)) {
                        try {
                            $DataDiskConfig = New-AzDiskConfig -SkuName $StorageType -Location $Location -CreateOption Empty -DiskSizeGB $DataDiskSize
                            Write-Log -LogFile $LogFile -LogString "Setting $DataDiskName config"
                        } catch {
                            Write-Log -LogFile $LogFile -LogString "ERROR setting $DataDiskName config : $_" -ForegroundColor Red
                        }
                        try {
                            $DataDisk = New-AzDisk -DiskName $DataDiskName -Disk $DataDiskConfig -ResourceGroupName $ResourceGroupName
                            Write-Log -LogFile $LogFile -LogString "Creating $DataDiskName"
                        } catch {
                            Write-Log -LogFile $LogFile -LogString "ERROR creating $DataDiskName : $_" -ForegroundColor Red
                        }
                        try {
                            $VirtualMachine = Add-AzVMDataDisk -VM $VirtualMachine -Name $DataDiskName -CreateOption Attach -ManagedDiskId $DataDisk.Id -Lun 0
                            Write-Log -LogFile $LogFile -LogString "Adding $DataDiskName to $VirtualMachineName"
                        } catch {
                            Write-Log -LogFile $LogFile -LogString "ERROR adding $DataDiskName to $VirtualMachineName : $_" -ForegroundColor Red
                        }
                    }
                }
            } else {
                try {
                    $VirtualMachine = Set-AzVMSourceImage -VM $VirtualMachine -PublisherName $ClientPublisherName -Offer $ClientOffer -Skus $ClientSkus -Version "latest"
                    Write-Log -LogFile $LogFile -LogString "Setting $VirtualMachineName source image"
                } catch {
                    Write-Log -LogFile $LogFile -LogString "ERROR setting $VirtualMachineName source image : $_" -ForegroundColor Red
                }
            }
            try {
                $VirtualMachine = Set-AzVMOSDisk -VM $VirtualMachine -Name $OSDiskName -StorageAccountType $StorageType -Caching ReadWrite -CreateOption FromImage
                Write-Log -LogFile $LogFile -LogString "Setting $VirtualMachineName OS disk"
            } catch {
                Write-Log -LogFile $LogFile -LogString "ERROR setting $VirtualMachineName OS disk : $_" -ForegroundColor Red
            }
            try {
                New-AzVM -VM $VirtualMachine -ResourceGroupName $ResourceGroupName -Location $location
                Write-Log -LogFile $LogFile -LogString "Creating $VirtualMachineName"
            } catch {
                Write-Log -LogFile $LogFile -LogString "ERROR creating $VirtualMachineName : $_" -ForegroundColor Red
            }
        } catch {
            Write-Log -LogFile $LogFile -LogString "ERROR creating $VirtualMachineName : $_" -ForegroundColor Red
            continue
        }
    }
}
try {
    # Clear sensitive variables securely
    Write-Log -LogFile $LogFile -LogString "Clearing credentials and disconnecting session"
    if ($VirtualMachineCredential) {
        $VirtualMachineCredential.Password.Dispose()
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
} finally {
    Disconnect-AzAccount -ErrorAction SilentlyContinue -confirm:$false | Out-Null
    Clear-AzContext -ErrorAction SilentlyContinue -force | Out-Null
}
