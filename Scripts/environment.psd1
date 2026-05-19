@{
    # ===================================================================
    # Corp-Account-Management environment configuration
    # ===================================================================
    # Loaded by Get-EnvironmentConfig in the corpadmin.psm1 module.
    # Place this file at the root of \Scripts\.
    # 
    # Per-deployment values (Domain NetBIOS, DnsSuffix, EmailSuffix) are
    # NOT in this file - they remain script parameters because they vary
    # per customer/tenant and the same scripts get reused across many.
    # ===================================================================
    
    # -------------------------------------------------------------------
    # Network topology
    # -------------------------------------------------------------------
    Network = @{
        SiteNames    = @('Site1', 'Site2')
        SiteSubnets  = @('10.71.104', '10.71.105')   # First three octets per site
        SiteNetmask  = "255.255.255.0"
        SiteCidr     = 24                            # Per-site subnet mask
        VnetPrefix   = 21                            # Azure VNet covers both sites
        DhcpStart    = 10                            # Last-octet start for DHCP scope
        DhcpEnd      = 254                           # Last-octet end for DHCP scope
        # Last-octet host offsets used inside each site subnet for static infra
        HostOffsetsLocal = @{
            RTR     = 1
            DC1     = 2
            DC2     = 3
            RootCA  = 4
            InterCA = 5
            WSUS    = 6
            EXCH    = 7
        }
        HostOffsetsAzure = @{
            DC1     = 4
        }
        # Network adapter naming (Local platform only)
        Interfaces = @{
            External = 'WAN'
            Internal = 'LAN'
        }
    }
    
    # -------------------------------------------------------------------
    # Active Directory OUs
    # -------------------------------------------------------------------
    OUs = @{
        # Top-level OUs (siblings of the domain root)
        Staff           = 'Staff'
        Administration  = 'Administration'
        Groups          = 'Groups'
        DomainComputers = 'Domain Computers'
        
        # Beneath OU=Administration
        HiPrivAccounts            = 'Hi_Priv_Accounts'
        HiPrivGroups              = 'Hi_Priv_Groups'
        LocalAdminGroups          = 'Local_Admin_Groups'
        ServiceAccounts           = 'Service_Accounts'
        SharedMailboxAccounts     = 'Shared_Mailbox_Accounts'
        EquipmentMailboxAccounts  = 'Equipment_Mailbox_Accounts'
        RoomMailboxAccounts       = 'Room_Mailbox_Accounts'
        
        # Beneath OU=Groups
        SharedMailboxAccess     = 'Shared_Mailbox_Access'
        EquipmentMailboxAccess  = 'Equipment_Mailbox_Access'
        RoomMailboxAccess       = 'Room_Mailbox_Access'
        
        # Beneath OU=Domain Computers
        Servers  = 'Servers'
        Desktops = 'Desktops'
        Laptops  = 'Laptops'
        VMs      = 'VMs'
    }
    
    # -------------------------------------------------------------------
    # Group names and naming prefixes
    # -------------------------------------------------------------------
    Groups = @{
        # Standing groups created by DomainSetup.ps1
        Staff       = 'Staff'
        IT          = 'IT'
        ITAdmin     = 'IT_Admin'
        O365License = 'License_Office365'
        TaskPrefix  = 'ADM_Task_'
        RolePrefix  = 'ADM_Role_'
        
        # Conventional prefixes for groups created at user/account creation time
        SharedAccessPrefix    = 'sh_'
        EquipmentAccessPrefix = 'eq_'
        RoomAccessPrefix      = 'ro_'
    }
    
    # -------------------------------------------------------------------
    # SMB / DFS shares
    # -------------------------------------------------------------------
    Shares = @{
        Root        = 'Store'
        Profiles    = 'Profiles'
        Software    = 'Software'
        WSUS        = 'WSUS'
        CertEnroll  = 'CertEnroll'
    }
    
    # -------------------------------------------------------------------
    # Regional / locale defaults applied to mailboxes and Entra users
    # -------------------------------------------------------------------
    Locale = @{
        Language      = 'en-GB'
        TimeZone      = 'GMT Standard Time'
        DateFormat    = 'dd/MM/yyyy'
        TimeFormat    = 'HH:mm'
        UsageLocation = 'GB'
        Dictionary    = 'EnglishUnitedKingdom'
    }
    
    # -------------------------------------------------------------------
    # Security defaults
    # -------------------------------------------------------------------
    Security = @{
        PasswordLength       = 20
        MaxElevationMinutes  = 480   # Cap for ElevateUser.ps1 -TimeSpan
    }
    
    # -------------------------------------------------------------------
    # Azure / azure_buildout.ps1
    # -------------------------------------------------------------------
    Azure = @{
        Location    = 'UK South'
        StorageType = 'StandardSSD_LRS'
        DataDiskGB  = 64
        ServerImage = @{
            Publisher = 'MicrosoftWindowsServer'
            Offer     = 'WindowsServer'
            Sku       = '2025-Datacenter-g2'
        }
        ClientImage = @{
            Publisher = 'MicrosoftWindowsDesktop'
            Offer     = 'windows-11'
            Sku       = 'win11-25h2-pro'
        }
        VmSizes = @{
            Default  = 'Standard_B2s'
            Exchange = 'Standard_D4s_v5'
            Client   = 'Standard_D2s_v5'
        }
        # Common tags - script merges in runtime Department/Owner values
        BaseTags = @{
            Criticality = 'Tier 1'
            Environment = 'Production'
        }
    }
    
    # -------------------------------------------------------------------
    # Exchange
    # -------------------------------------------------------------------
    Exchange = @{
        Subdomain = 'exchange'   # Renders as exchange.<EmailSuffix>
    }
    
    # -------------------------------------------------------------------
    # Entra ID directory roles by privilege level
    # Used by CreateITCloudAdminUser.ps1
    # -------------------------------------------------------------------
    EntraRoles = @{
        Level1 = @(
            'Helpdesk Administrator'
            'Service Support Administrator'
            'Global Reader'
        )
        Level2 = @(
            'User Administrator'
            'Groups Administrator'
            'Authentication Administrator'
            'License Administrator'
        )
        Level3 = @(
            'Exchange Administrator'
            'Teams Administrator'
            'SharePoint Administrator'
            'Privileged Authentication Administrator'
            'Privileged Role Administrator'
        )
    }
    
    # -------------------------------------------------------------------
    # WSUS sync configuration
    # -------------------------------------------------------------------
    WSUS = @{
        Products = @(
            'Windows 11'
            'Windows Server 2022'
            'Windows Server 2025'
        )
        Classifications = @(
            'Critical Updates'
            'Definition Updates'
            'Feature Packs'
            'Security Updates'
            'Service Packs'
            'Update Rollups'
            'Updates'
        )
    }
}
