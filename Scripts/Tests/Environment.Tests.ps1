#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.5.0' }

BeforeAll {
    $script:scriptsRoot = Join-Path $PSScriptRoot '..'
    $script:envConfig   = Import-PowerShellDataFile (Join-Path $script:scriptsRoot 'environment.psd1')
}

Describe 'environment.psd1 structure' {
    It 'has all top-level sections' {
        $expected = 'Network','OUs','Groups','Shares','Locale','Security','Azure','Exchange','EntraRoles','WSUS'
        foreach ($k in $expected) { $script:envConfig.Keys | Should -Contain $k }
    }
    It 'Groups hash has required prefix and access keys' {
        foreach ($key in @(
            'TaskPrefix','RolePrefix','Staff','IT','ITAdmin','O365License','SharedAccessPrefix','EquipmentAccessPrefix','RoomAccessPrefix'
        )) {
            $script:envConfig.Groups.Keys | Should -Contain $key
        }
    }
    It 'OUs hash has all required keys' {
        foreach ($key in @('Staff','Administration','Groups','DomainComputers','HiPrivAccounts','HiPrivGroups'
            ,'SharedMailboxAccounts','EquipmentMailboxAccounts','RoomMailboxAccounts'
            ,'SharedMailboxAccess','EquipmentMailboxAccess','RoomMailboxAccess'
            ,'LocalAdminGroups','ServiceAccounts','Servers','Desktops','Laptops','VMs'
        )) {
            $script:envConfig.OUs.Keys | Should -Contain $key
        }
    }
    It 'Network.HostOffsetsLocal has all required role keys' {
        foreach ($key in @('RTR','DC1','DC2','RootCA','InterCA','WSUS','EXCH')) {
            $script:envConfig.Network.HostOffsetsLocal.Keys | Should -Contain $key
        }
    }
    It 'EntraRoles has Level1/2/3 as arrays' {
        $script:envConfig.EntraRoles.Level1 -is [array] | Should -BeTrue
        $script:envConfig.EntraRoles.Level2 -is [array] | Should -BeTrue
        $script:envConfig.EntraRoles.Level3 -is [array] | Should -BeTrue
    }
    It 'Network.SiteSubnets length matches Network.SiteNames length' {
        $script:envConfig.Network.SiteSubnets.Count | Should -Be $script:envConfig.Network.SiteNames.Count
    }
    It 'Security.MaxElevationMinutes is a positive integer' {
        $script:envConfig.Security.MaxElevationMinutes | Should -BeOfType [int]
        $script:envConfig.Security.MaxElevationMinutes | Should -BeGreaterThan 0
        $script:envConfig.Security.MaxElevationMinutes | Should -BeLessOrEqual 480
    }
    It 'Security.PasswordLength is within New-Password ValidateRange (12..256)' {
        $script:envConfig.Security.PasswordLength | Should -BeGreaterOrEqual 12
        $script:envConfig.Security.PasswordLength | Should -BeLessOrEqual 256
    }
    It 'Locale.UsageLocation is a two-letter ISO country code' {
        $script:envConfig.Locale.UsageLocation | Should -Match '^[A-Z]{2}$'
    }
}

# =============================================================================
# Network value pinning
# =============================================================================
Describe 'environment.psd1 Network values' {
    # The two SiteSubnets are the network identity of the Corp deployment.
    # Drift here breaks DHCP scopes, AD site/subnet definitions, and the
    # static IP assignments computed by setup.ps1 -Role.
    It 'SiteSubnets[0] is 10.71.104' {
        $script:envConfig.Network.SiteSubnets[0] | Should -Be '10.71.104'
    }
    It 'SiteSubnets[1] is 10.71.105' {
        $script:envConfig.Network.SiteSubnets[1] | Should -Be '10.71.105'
    }
    It 'SiteNames are Site1 and Site2' {
        $script:envConfig.Network.SiteNames | Should -Be @('Site1','Site2')
    }
    It 'SiteCidr is /24 (per-site mask)' {
        $script:envConfig.Network.SiteCidr | Should -Be 24
    }
    It 'VnetPrefix is /21 (covers both sites)' {
        # Azure VNet sits one notch wider than the per-site /24s so both
        # subnets can co-exist inside it.
        $script:envConfig.Network.VnetPrefix | Should -Be 21
    }
    It 'DhcpStart is 10 (preserves .1-.9 for static infra)' {
        $script:envConfig.Network.DhcpStart | Should -Be 10
    }
    It 'DhcpEnd is 254' {
        $script:envConfig.Network.DhcpEnd | Should -Be 254
    }
    It 'DhcpStart sits above all HostOffsetsLocal values' {
        # Static infra (RTR/DC1/DC2/RootCA/InterCA/WSUS/EXCH) occupies
        # low-numbered host offsets. DhcpStart must be strictly above
        # the highest static offset, otherwise DHCP would hand out the
        # same address as a static infra box.
        $maxStatic = ($script:envConfig.Network.HostOffsetsLocal.Values | Measure-Object -Maximum).Maximum
        $script:envConfig.Network.DhcpStart | Should -BeGreaterThan $maxStatic
    }
    It 'Interfaces.External is WAN' {
        # dhcp.ps1 expects these exact names when renaming NICs.
        $script:envConfig.Network.Interfaces.External | Should -Be 'WAN'
    }
    It 'Interfaces.Internal is LAN' {
        $script:envConfig.Network.Interfaces.Internal | Should -Be 'LAN'
    }
}

# =============================================================================
# Locale value pinning (UK deployment)
# =============================================================================
Describe 'environment.psd1 Locale values (UK)' {
    # Cross-checked against CorpAdmin.Mailbox.Tests.ps1 source-level guards
    # which forbid en-US / Pacific Standard Time / EnglishUnitedStates
    # anywhere in the module. These pin the positive side of that contract.
    It 'Language is en-GB' {
        $script:envConfig.Locale.Language | Should -Be 'en-GB'
    }
    It 'TimeZone is GMT Standard Time' {
        $script:envConfig.Locale.TimeZone | Should -Be 'GMT Standard Time'
    }
    It 'DateFormat is dd/MM/yyyy' {
        $script:envConfig.Locale.DateFormat | Should -Be 'dd/MM/yyyy'
    }
    It 'TimeFormat is 24-hour HH:mm' {
        $script:envConfig.Locale.TimeFormat | Should -Be 'HH:mm'
    }
    It 'UsageLocation is GB' {
        $script:envConfig.Locale.UsageLocation | Should -Be 'GB'
    }
    It 'Dictionary is EnglishUnitedKingdom' {
        $script:envConfig.Locale.Dictionary | Should -Be 'EnglishUnitedKingdom'
    }
}

# =============================================================================
# Security value pinning
# =============================================================================
Describe 'environment.psd1 Security values' {
    # MaxElevationMinutes is the CAP applied by ElevateUser.ps1 and
    # Enable-CloudAdmin.ps1. Reducing it tightens the security stance;
    # increasing it loosens it. Either direction should be a deliberate
    # decision, not an accidental edit.
    It 'MaxElevationMinutes is 480 (8 hours)' {
        $script:envConfig.Security.MaxElevationMinutes | Should -Be 480
    }
    It 'PasswordLength is 20' {
        # Matches the New-Password default; changing one without the
        # other produces inconsistent generation vs validation.
        $script:envConfig.Security.PasswordLength | Should -Be 20
    }
}

# =============================================================================
# Azure value pinning
# =============================================================================
Describe 'environment.psd1 Azure values' {
    It 'Location is UK South' {
        $script:envConfig.Azure.Location | Should -Be 'UK South'
    }
    It 'StorageType is StandardSSD_LRS' {
        # Premium SSD would inflate cost; HDD would slow VM boot. Standard
        # SSD is the deliberate middle ground.
        $script:envConfig.Azure.StorageType | Should -Be 'StandardSSD_LRS'
    }
    It 'DataDiskGB is 64' {
        $script:envConfig.Azure.DataDiskGB | Should -Be 64
    }
    It 'ServerImage publisher is MicrosoftWindowsServer' {
        $script:envConfig.Azure.ServerImage.Publisher | Should -Be 'MicrosoftWindowsServer'
    }
    It 'ServerImage offer is WindowsServer' {
        $script:envConfig.Azure.ServerImage.Offer | Should -Be 'WindowsServer'
    }
    It 'ServerImage SKU is 2025-Datacenter-g2' {
        $script:envConfig.Azure.ServerImage.Sku | Should -Be '2025-Datacenter-g2'
    }
    It 'ClientImage publisher is MicrosoftWindowsDesktop' {
        $script:envConfig.Azure.ClientImage.Publisher | Should -Be 'MicrosoftWindowsDesktop'
    }
    It 'ClientImage offer is windows-11' {
        $script:envConfig.Azure.ClientImage.Offer | Should -Be 'windows-11'
    }
    It 'ClientImage SKU matches win11-<release>-pro pattern' {
        # The exact release name is forward-pointing; pin the shape rather
        # than the exact string so 25H2 -> 25H3 etc. doesn't trip the test,
        # but a typo or accidental non-pro SKU does.
        $script:envConfig.Azure.ClientImage.Sku | Should -Match '^win11-\d{2}h\d-pro$'
    }
    It 'VmSizes.Default is a B-series burstable SKU' {
        # B-series for the cheap baseline VMs is deliberate.
        $script:envConfig.Azure.VmSizes.Default | Should -Match '^Standard_B'
    }
    It 'VmSizes.Exchange is Standard_D4s_v5 (4 vCPU / 16 GB RAM)' {
        # Exchange has a hard 8 GB minimum and is unhappy below D4s.
        $script:envConfig.Azure.VmSizes.Exchange | Should -Be 'Standard_D4s_v5'
    }
    It 'VmSizes.Client is Standard_D2s_v5' {
        $script:envConfig.Azure.VmSizes.Client | Should -Be 'Standard_D2s_v5'
    }
    It 'BaseTags includes Criticality and Environment' {
        $script:envConfig.Azure.BaseTags.Keys | Should -Contain 'Criticality'
        $script:envConfig.Azure.BaseTags.Keys | Should -Contain 'Environment'
    }
}

# =============================================================================
# Groups value pinning
# =============================================================================
Describe 'environment.psd1 Groups values' {
    # Drift on these prefixes silently breaks every script that constructs
    # group names by concatenation (DomainSetup, CreateITAdminUser,
    # CreateGroup, ElevateUser, etc.).
    It "TaskPrefix is 'ADM_Task_'" {
        $script:envConfig.Groups.TaskPrefix | Should -Be 'ADM_Task_'
    }
    It "RolePrefix is 'ADM_Role_'" {
        $script:envConfig.Groups.RolePrefix | Should -Be 'ADM_Role_'
    }
    It "Staff is 'Staff'" {
        $script:envConfig.Groups.Staff | Should -Be 'Staff'
    }
    It "IT is 'IT'" {
        $script:envConfig.Groups.IT | Should -Be 'IT'
    }
    It "ITAdmin is 'IT_Admin'" {
        $script:envConfig.Groups.ITAdmin | Should -Be 'IT_Admin'
    }
    It "SharedAccessPrefix is 'sh_'" {
        # Pinned by CorpAdmin.Mailbox.Tests.ps1 too - keep in lockstep.
        $script:envConfig.Groups.SharedAccessPrefix | Should -Be 'sh_'
    }
    It "EquipmentAccessPrefix is 'eq_'" {
        $script:envConfig.Groups.EquipmentAccessPrefix | Should -Be 'eq_'
    }
    It "RoomAccessPrefix is 'ro_'" {
        $script:envConfig.Groups.RoomAccessPrefix | Should -Be 'ro_'
    }
}

# =============================================================================
# Shares value pinning
# =============================================================================
Describe 'environment.psd1 Shares values' {
    # Share names are referenced by DomainSetup.ps1 for ACL creation,
    # by additional-DFS-Setup.ps1 for DFS replication groups, and by
    # WSUS / cert-enrollment scripts.
    It "Root share is 'Store'" {
        $script:envConfig.Shares.Root | Should -Be 'Store'
    }
    It "Profiles share is 'Profiles'" {
        $script:envConfig.Shares.Profiles | Should -Be 'Profiles'
    }
    It "Software share is 'Software'" {
        $script:envConfig.Shares.Software | Should -Be 'Software'
    }
    It "WSUS share is 'WSUS'" {
        $script:envConfig.Shares.WSUS | Should -Be 'WSUS'
    }
    It "CertEnroll share is 'CertEnroll'" {
        $script:envConfig.Shares.CertEnroll | Should -Be 'CertEnroll'
    }
}

# =============================================================================
# Exchange / EntraRoles value pinning
# =============================================================================
Describe 'environment.psd1 Exchange values' {
    It "Subdomain is 'exchange' (renders as exchange.<EmailSuffix>)" {
        $script:envConfig.Exchange.Subdomain | Should -Be 'exchange'
    }
}

Describe 'environment.psd1 EntraRoles tiering' {
    # The tiering must escalate: Level1 > Level1+Level2 > Level1+Level2+Level3.
    # Drift here (e.g. moving 'Global Reader' to Level3) would silently
    # widen or narrow what each privilege level grants.
    It 'Level1 includes Helpdesk Administrator (Tier-1 desktop admins)' {
        $script:envConfig.EntraRoles.Level1 | Should -Contain 'Helpdesk Administrator'
    }
    It 'Level1 includes Global Reader (read-only directory access)' {
        $script:envConfig.EntraRoles.Level1 | Should -Contain 'Global Reader'
    }
    It 'Level2 includes User Administrator' {
        $script:envConfig.EntraRoles.Level2 | Should -Contain 'User Administrator'
    }
    It 'Level2 includes Authentication Administrator' {
        $script:envConfig.EntraRoles.Level2 | Should -Contain 'Authentication Administrator'
    }
    It 'Level3 includes Privileged Authentication Administrator' {
        # The escalation gate - having PAA means you can reset MFA for
        # other admins. Level3 only.
        $script:envConfig.EntraRoles.Level3 | Should -Contain 'Privileged Authentication Administrator'
    }
    It 'Level3 includes Privileged Role Administrator' {
        $script:envConfig.EntraRoles.Level3 | Should -Contain 'Privileged Role Administrator'
    }
    It 'no role appears at more than one level (de-duplication)' {
        $all = @($script:envConfig.EntraRoles.Level1) +
               @($script:envConfig.EntraRoles.Level2) +
               @($script:envConfig.EntraRoles.Level3)
        $dupes = $all | Group-Object | Where-Object Count -gt 1
        $dupes | Should -BeNullOrEmpty
    }
}

# =============================================================================
# WSUS value pinning
# =============================================================================
Describe 'environment.psd1 WSUS values' {
    # WSUS Products spans both server versions by design: WDS builds Server 2022
    # on-prem, azure_buildout builds Server 2025, plus Windows 11 clients. The
    # breadth is the union of the two build paths, not config drift.
    It 'syncs Windows 11' {
        $script:envConfig.WSUS.Products | Should -Contain 'Windows 11'
    }
    It 'syncs Windows Server 2022' {
        $script:envConfig.WSUS.Products | Should -Contain 'Windows Server 2022'
    }
    It 'syncs Windows Server 2025' {
        $script:envConfig.WSUS.Products | Should -Contain 'Windows Server 2025'
    }
    It 'classifications include Security Updates and Critical Updates' {
        $script:envConfig.WSUS.Classifications | Should -Contain 'Security Updates'
        $script:envConfig.WSUS.Classifications | Should -Contain 'Critical Updates'
    }
    It 'classifications do NOT include Drivers (avoids driver-update churn)' {
        # If you ever DO want WSUS to manage drivers, this test is the
        # safety net to make the decision conscious.
        $script:envConfig.WSUS.Classifications | Should -Not -Contain 'Drivers'
    }
}
