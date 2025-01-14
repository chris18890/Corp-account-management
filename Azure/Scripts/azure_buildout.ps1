[CmdletBinding()]
param(
    [string]$SubscriptionId
    , [string]$TenantID
    , [string]$MyIPAddress
    , [string]$NameStem
)
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
$IPNet = @("10.71.104", "10.71.105")
$Site = @("Site1", "Site2")
$Subnet = @()
$SubnetCDIR = "24"
$VnetCDIR = "21"
$DNSServer4 = "4"
$VnetDNSServer = $IPNet[0] + "." + $DNSServer4
$VnetPrefix = $IPNet[0] + ".0/" + $VnetCDIR
Connect-AzAccount -SubscriptionId "$SubscriptionId" -TenantID "$TenantID"
Set-AzContext -SubscriptionId "$SubscriptionId" -TenantID "$TenantID"
Write-Verbose -Message ("{0} - {1}" -f (Get-Date).ToString(),"Verifying the Azure login subscription status...")
if(-not $(Get-AzContext)) {
    Write-Verbose -Message ("{0} - {1}" -f (Get-Date).ToString(),"Login to Azure subscription failed, no valid subscription found.")
    Connect-AzAccount -SubscriptionId "$SubscriptionId" -TenantID "$TenantID"
    Set-AzContext -SubscriptionId "$SubscriptionId" -TenantID "$TenantID"
}
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
$ClientSkus = "win11-24h2-pro"
$MachineNames = @("$NameStem-DC1", "$NameStem-DC2", "$NameStem-RTR", "$NameStem-EXCH", "$NameStem-1")
$VirtualMachineCredential = Get-Credential -Message 'Please enter the vm credentials'
$BDStorageAccountName = $NameStemLower + "infrargbootdiag"
$BDStorageAccountSKU = "Standard_GRS"
$MachineSKU = "Standard_B2s"
$ExchangeMachineSKU = "Standard_D4s_v3"
$ClientMachineSKU = "Standard_D2s_v3"

$ResourceGroup = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
if (!($ResourceGroup)) {
    Write-Verbose -Message ("{0} - {1}" -f (Get-Date).ToString(),"Creating new resource group $ResourceGroupName")
    $ResourceGroup = New-AzResourceGroup -Name $ResourceGroupName -Location $Location -Tag @{Empty=$null; Criticality = "Tier 1"; Department = "$NameStem"; Environment = "Production"; Owner = "Chris Murray"}
    Write-Verbose -Message ("{0} - {1}" -f (Get-Date).ToString(),"New resource group $ResourceGroupName is created.")
} else {
    Write-Verbose -Message ("{0} - {1}" -f (Get-Date).ToString(),"$ResourceGroupName Resource Group already exists, skipping new resource group creation...")
}

$VirtualNetwork = Get-AzVirtualNetwork -Name $VirtualNetworkName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
if (!($VirtualNetwork)) {
    Write-Verbose -Message ("{0} - {1}" -f (Get-Date).ToString(),"Creating new virtual network $VirtualNetworkName")
    $SecurityGroup = Get-AzNetworkSecurityGroup -Name $SecurityGroupName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
    if (!($SecurityGroup)) {
        $nsgRuleRDP = New-AzNetworkSecurityRuleConfig -Name rdp-rule -Description "Allow RDP" -Access Allow -Protocol Tcp -Direction Inbound -Priority 100 -SourceAddressPrefix $MyIPAddress -SourcePortRange * -DestinationAddressPrefix * -DestinationPortRange 3389
        $SecurityGroup = New-AzNetworkSecurityGroup -Name $SecurityGroupName -ResourceGroupName $ResourceGroupName -Location $Location -SecurityRules $nsgRuleRDP
    }
    For ($i=0; $i -le $IPNet.Length -1; $i++){
        $ScopeNetwork = $IPNet[$i] + ".0"
        $ScopeName = "$ScopeNetwork/$SubnetCDIR"
        $Subnet += New-AzVirtualNetworkSubnetConfig -Name $Site[$i] -AddressPrefix $ScopeName -NetworkSecurityGroup $SecurityGroup
    }
    $VirtualNetwork = New-AzVirtualNetwork -Name $VirtualNetworkName -ResourceGroupName $ResourceGroupName -Location $Location -AddressPrefix $VnetPrefix -Subnet ($Subnet) -DnsServer $VnetDNSServer
    Write-Verbose -Message ("{0} - {1}" -f (Get-Date).ToString(),"New virtual network $VirtualNetworkName is created.")
} else {
    Write-Verbose -Message ("{0} - {1}" -f (Get-Date).ToString(),"$VirtualNetworkName virtual network already exists, skipping new virtual network creation...")
}

$BDStorageAccount = Get-AzStorageAccount -Name $BDStorageAccountName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
if (!($BDStorageAccount)) {
    Write-Verbose -Message ("{0} - {1}" -f (Get-Date).ToString(),"Creating new Storage Account $BDStorageAccountName")
    $BDStorageAccount = New-AzStorageAccount -ResourceGroupName $ResourceGroupName -AccountName $BDStorageAccountName -Location $Location -SkuName $BDStorageAccountSKU
    Write-Verbose -Message ("{0} - {1}" -f (Get-Date).ToString(),"New Storage Account $BDStorageAccountName is created.")
} else {
    Write-Verbose -Message ("{0} - {1}" -f (Get-Date).ToString(),"$BDStorageAccountName Storage Account already exists, skipping new storage account creation...")
}

foreach ($Machine in $MachineNames){
    $VirtualMachineName = $Machine
    $VM = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $VirtualMachineName -ErrorAction SilentlyContinue
    if (!($VM)) {
        $OSDiskName = $VirtualMachineName + "_OSDisk"
        $DataDiskName = $VirtualMachineName + "_DataDisk_0"
        $VirtualMachineNICName = $VirtualMachineName + "_NIC"
        $VirtualMachineSize = $MachineSKU
        $VirtualMachineNIC = Get-AzNetworkInterface -Name $VirtualMachineNICName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
        if (!($VirtualMachineNIC)) {
            $VirtualMachineNIC = New-AzNetworkInterface -Name $VirtualMachineNICName -ResourceGroupName $ResourceGroupName -Location $Location -SubnetId $VirtualNetwork.Subnets[0].Id
        }
        if ($VirtualMachineName -eq "$NameStem-EXCH") {
            $VirtualMachineSize = $ExchangeMachineSKU
            $PublicIPAddressName = $VirtualMachineName + "_PIP"
            $PublicIPAddress = Get-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
            if (!($PublicIPAddress)) {
                $PublicIPAddress = New-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName -Location $Location -AllocationMethod Static
            }
            $VirtualMachineNIC | Set-AzNetworkInterfaceIpConfig -Name $VirtualMachineNIC.IpConfigurations[0].Name -PublicIpAddressId $PublicIPAddress.Id | Set-AzNetworkInterface
            $VirtualMachine_PublicIP = (Get-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName).IpAddress
            $VirtualMachine_PrivateIP = (Get-AzNetworkInterface -Name $VirtualMachineNICName -ResourceGroupName $ResourceGroupName).IpConfigurations.PrivateIpAddress
            $VirtualMachineNIC_IP = @("$VirtualMachine_PrivateIP","$VirtualMachine_PublicIP")
            $SecurityGroup = Get-AzNetworkSecurityGroup -Name $SecurityGroupName -ResourceGroupName $ResourceGroupName
            if (!(Get-AzNetworkSecurityRuleConfig -Name http-rule -NetworkSecurityGroup $SecurityGroup)) {
                $SecurityGroup | Add-AzNetworkSecurityRuleConfig -Name http-rule -Description "Allow HTTP" -Access Allow -Protocol Tcp -Direction Inbound -Priority 110 -SourceAddressPrefix * -SourcePortRange * -DestinationAddressPrefix $VirtualMachineNIC_IP -DestinationPortRange 80
                $SecurityGroup | Set-AzNetworkSecurityGroup
            }
            if (!(Get-AzNetworkSecurityRuleConfig -Name https-rule -NetworkSecurityGroup $SecurityGroup)) {
                $SecurityGroup | Add-AzNetworkSecurityRuleConfig -Name https-rule -Description "Allow HTTPS" -Access Allow -Protocol Tcp -Direction Inbound -Priority 120 -SourceAddressPrefix * -SourcePortRange * -DestinationAddressPrefix $VirtualMachineNIC_IP -DestinationPortRange 443
                $SecurityGroup | Set-AzNetworkSecurityGroup
            }
        }
        if ($VirtualMachineName -eq "$NameStem-1") {
            $VirtualMachineSize = $ClientMachineSKU
            $PublicIPAddressName = $VirtualMachineName + "_PIP"
            $PublicIPAddress = Get-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
            if (!($PublicIPAddress)) {
                $PublicIPAddress = New-AzPublicIpAddress -Name $PublicIPAddressName -ResourceGroupName $ResourceGroupName -Location $Location -AllocationMethod Static
            }
            $VirtualMachineNIC | Set-AzNetworkInterfaceIpConfig -Name $VirtualMachineNIC.IpConfigurations[0].Name -PublicIpAddressId $PublicIPAddress.Id | Set-AzNetworkInterface
        }
        $VirtualMachine = New-AzVMConfig -VMName $VirtualMachineName -VMSize $virtualMachineSize -Tags @{Criticality = "Tier 1"; Environment = "Production"}
        $VirtualMachine = Set-AzVMOperatingSystem -VM $VirtualMachine -Windows -ComputerName $VirtualMachineName -Credential $VirtualMachineCredential -ProvisionVMAgent -EnableAutoUpdate
        $VirtualMachine = Add-AzVMNetworkInterface -VM $VirtualMachine -Id $VirtualMachineNIC.Id
        $VirtualMachine = Set-AzVMBootDiagnostic -VM $VirtualMachine -Enable -ResourceGroupName $ResourceGroupName -StorageAccountName $BDStorageAccountName
        if ($VirtualMachineName -eq "$NameStem-DC1" -or $VirtualMachineName -eq "$NameStem-DC2" -or $VirtualMachineName -eq "$NameStem-RTR" -or $VirtualMachineName -eq "$NameStem-EXCH") {
            $VirtualMachine = Set-AzVMSourceImage -VM $VirtualMachine -PublisherName $ServerPublisherName -Offer $ServerOffer -Skus $ServerSkus -Version "latest"
            if ($VirtualMachineName -eq "$NameStem-DC1" -or $VirtualMachineName -eq "$NameStem-DC2") {
                $DataDisk = Get-AzDisk -Name $DataDiskName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
                if (!($DataDisk)) {
                    $DataDiskConfig = New-AzDiskConfig -SkuName $StorageType -Location $Location -CreateOption Empty -DiskSizeGB 64
                    $DataDisk = New-AzDisk -DiskName $DataDiskName -Disk $DataDiskConfig -ResourceGroupName $ResourceGroupName
                }
                $VirtualMachine = Add-AzVMDataDisk -VM $VirtualMachine -Name $DataDiskName -CreateOption Attach -ManagedDiskId $DataDisk.Id -Lun 0
            }
        } else {
            $VirtualMachine = Set-AzVMSourceImage -VM $VirtualMachine -PublisherName $ClientPublisherName -Offer $ClientOffer -Skus $ClientSkus -Version "latest"
        }
        $VirtualMachine = Set-AzVMOSDisk -VM $VirtualMachine -Name $OSDiskName -StorageAccountType $StorageType -Caching ReadWrite -CreateOption FromImage
        New-AzVM -VM $VirtualMachine -ResourceGroupName $ResourceGroupName -Location $location
    }
}
