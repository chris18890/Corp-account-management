#Requires -Modules ActiveDirectory, GroupPolicy
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

# Execution Tier: Tier-0

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
    ,[Parameter(Mandatory)][string]$Drive
    ,[string]$LogFile
)

#====================================================================
# Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$ServerName = "$env:computername"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$rootdse = Get-ADRootDSE
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$ParentOU = "Domain Computers"
$Location = "OU=$ParentOU,$EndPath"
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$SID500 = (Get-ADUser -Filter * -Server $DCHostName | Select-Object -Property SID,Name | Where-Object -Property SID -like "*-500").Name
$GPOLocation = Join-Path $PSScriptRoot "GPOs"
$ImportedGPOs = @()
$UserScriptsLocation = Join-Path (Split-Path $PSScriptRoot -Parent) "Users"
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptTitle = "$Domain Bootstrap Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_domain_setup_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}
#====================================================================

#====================================================================
# Drive where all the folders will be created
#====================================================================
$Drive = $Drive.TrimEnd(':') + ':'
$RootShare = "Store"
#====================================================================

#====================================================================
# Group Variables
#====================================================================
$GroupsOU = "Groups"
$GroupCategory = "Security"
$GroupScope = "Universal"
$StaffGroup = "Staff"
$AdministrationOU = "Administration"
$ITGroup = "IT"
$ITAdminGroup = "IT_Admin"
$O365LicenseGroup = "License_Office365"
$DNSOperatorsGroup = "ADM_Task_DNS_Operators"
$DNSReadOnlyGroup = "ADM_Task_DNS_ReadOnly"
$HiPrivAccountAdminGroup = "ADM_Task_HiPriv_Account_Admins"
$HiPrivGroupAdminGroup = "ADM_Task_HiPriv_Group_Admins"
$InstallerGroup = "ADM_Task_Installers"
$LocalAdminGroupAdminGroup = "ADM_Task_Local_Admin_Group_Admins"
$UserPasswordDelegationGroup = "ADM_Task_Password_Admins"
$SERAccessAdminGroup = "ADM_Task_SER_Access_Admins"
$SERAccountAdminGroup = "ADM_Task_SER_Account_Admins"
$ServiceAccountAdminGroup = "ADM_Task_Service_Account_Admins"
$StandardAccountAdminGroup = "ADM_Task_Standard_Account_Admins"
$StandardGroupAdminGroup = "ADM_Task_Standard_Group_Admins"
$EquipmentAccountsOU = "Equipment_Mailbox_Accounts,OU=$AdministrationOU"
$RoomAccountsOU = "Room_Mailbox_Accounts,OU=$AdministrationOU"
$SharedAccountsOU = "Shared_Mailbox_Accounts,OU=$AdministrationOU"
$EquipmentGroupsOU = "Equipment_Mailbox_Access,OU=$GroupsOU"
$RoomGroupsOU = "Room_Mailbox_Access,OU=$GroupsOU"
$SharedGroupsOU = "Shared_Mailbox_Access,OU=$GroupsOU"
#====================================================================

#====================================================================
# Set up logging
#====================================================================
function Write-Log {
    param([string]$LogString,[string]$ForegroundColor)
    #================================================================
    # Purpose:          To write a string with a date and time stamp to a log file
    # Assumptions:      $LogFile set with path to log file to write to
    # Effects:
    # Inputs:
    # $LogString:       String to write to log file
    # Calls:
    # Returns:
    # Notes:
    #================================================================
    "$(Get-Date -Format 'G') $LogString" | Out-File -Filepath $LogFile -Append -Encoding UTF8
    if ($ForegroundColor) {
        Write-Host $LogString -ForegroundColor $ForegroundColor
    } else {
        Write-Host $LogString
    }
}
#====================================================================

#====================================================================
# OU creation function
#====================================================================
function New-ADOU {
    [CmdletBinding()]
    param(
        [string]$OUName,[String]$Path,[String]$OUDescription
    )
    Write-Log "Creating OU $OUName"
    try {
        New-ADOrganizationalUnit -Name $OUName -Path $Path -ProtectedFromAccidentalDeletion:$true -Description $OUDescription -Server $DCHostName
        Write-Log "Created OU $OUName"
    } catch {
        $ex = $_.Exception
        if ($ex.Message -match "already exists") {
            Write-Log "'$OUName' already exists" -ForegroundColor Green
        } else {
            throw
        }
    }
}
#====================================================================

#====================================================================
# Group creation function
#====================================================================
function New-DomainGroup {
    [CmdletBinding()]
    param(
        [String]$GroupName,[String]$GroupScope,[ValidateSet("E","H","N")][String]$O365,[Boolean]$HiddenFromAddressListsEnabled,[String]$Path,[String]$GroupDescription
    )
    Write-Log "Creating Group $GroupName"
    try {
        New-ADGroup -GroupCategory $GroupCategory -GroupScope $GroupScope -Name $GroupName -Path $Path -SamAccountName $GroupName -Server $DCHostName -Description $GroupDescription
        Set-ADObject -Identity "CN=$GroupName,$Path" -Server $DCHostName -ProtectedFromAccidentalDeletion $true
        Write-Log "Created $GroupName"
    } catch {
        $ex = $_.Exception
        if ($ex.Message -match "already exists") {
            Write-Log "'$GroupName' already exists" -ForegroundColor Green
        } else {
            throw
        }
    }
    if ($O365 -eq "E" -or $O365 -eq "H") {
        try {
            Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
            Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $HiddenFromAddressListsEnabled -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
        } catch {
            Write-Log "WARNING: Could not enable $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}
#====================================================================

#====================================================================
# Group addition function
#====================================================================
function Add-GroupMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Group
        , [Parameter(Mandatory)][string]$Member
    )
    #================================================================
    # Purpose:          To add a user account or group to a group
    # Assumptions:      Parameters have been set correctly
    # Effects:          Member will be added to the group
    # Inputs:           $Group - Group name as set before calling the function
    #                   $Member - Object to be added
    # Calls:            Write-Log function
    # Returns:
    # Notes:
    #================================================================
    $checkGroup = Get-ADGroup -LDAPFilter "(SAMAccountName=$Group)" -Server $DCHostName
    if ($null -ne $checkGroup) {
        $checkMember = Get-ADObject -LDAPFilter "(SAMAccountName=$Member)" -Server $DCHostName
        if (-not $checkMember) {
            Write-Log "'$Member' does not exist" -ForegroundColor Red
            return
        }
        Write-Log "Adding $Member to $Group"
        try {
            Add-ADGroupMember -Identity $Group -Members $Member -Server $DCHostName
            Write-Log "Added $Member to $Group"
        } catch {
            $ex = $_.Exception
            if ($ex.Message -match "already a member") {
                Write-Log "'$Member' is already a member of group '$Group'" -ForegroundColor Green
            } else {
                throw
            }
        }
    } else {
        Write-Log "$Group does not exist" -ForegroundColor Red
    }
}
#====================================================================

#====================================================================
# GPO link function
#====================================================================
function Add-GPOLink {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$GPOName,
        [Parameter(Mandatory)][string]$GPOTarget
    )
    #================================================================
    # Purpose:          To link a GPO to an OU
    # Assumptions:      Parameters have been set correctly
    # Effects:          GPO will be linked to the OU
    # Inputs:           $GPOName - Name of GPO as set before calling the function
    #                   $GPOTarget - OU where GPO will be linked
    # Returns:
    # Notes:
    #================================================================
    Write-Log "Linking $GPOName to $GPOTarget"
    try {
        New-GPLink -name $GPOName -target $GPOTarget -LinkEnabled Yes -enforced yes -Order 1 -ErrorAction Stop -Server $DCHostName
        Write-Log "Linked $GPOName to $GPOTarget"
    } catch {
        $ex = $_.Exception
        if ($ex.Message -match "already exists") {
            Write-Log "'$GPOName' already linked to $GPOTarget" -ForegroundColor Green
        } else {
            throw
        }
    }
}
#====================================================================

#====================================================================
# Delegate permission on computer objects to a group
#====================================================================
function Grant-ComputerJoinDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"CreateChild",$AccessControlTypeAllow,$GuidMap["computer"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"DeleteChild",$AccessControlTypeAllow,$GuidMap["computer"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,"Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteDacl",$AccessControlTypeAllow,"Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"Self",$AccessControlTypeAllow,$ExtendedRightsMap["Validated write to DNS host name"],"Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"Self",$AccessControlTypeAllow,$ExtendedRightsMap["Validated write to service principal name"],"Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ExtendedRightAccessRule $AdminGroupSID,$AccessControlTypeAllow,$ExtendedRightsMap["Reset Password"],"Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ExtendedRightAccessRule $AdminGroupSID,$AccessControlTypeAllow,$ExtendedRightsMap["Change Password"],"Descendents",$GuidMap["computer"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permission on group objects to a group
#====================================================================
function Grant-GroupDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"CreateChild",$AccessControlTypeAllow,$GuidMap["group"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"DeleteChild",$AccessControlTypeAllow,$GuidMap["group"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,"Descendents",$GuidMap["group"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate modify membership permission on group objects to a group
#====================================================================
function Grant-GroupMembershipEditDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,$GuidMap["member"],'Descendents',$GuidMap["group"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate change password permission on user objects to a group
#====================================================================
function Grant-PasswordResetDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,$GuidMap["pwdLastSet"],'Descendents',$GuidMap["user"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,$GuidMap["lockoutTime"],'Descendents',$GuidMap["user"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"ExtendedRight",$AccessControlTypeAllow,$ExtendedRightsMap["Reset Password"],'Descendents',$GuidMap["user"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permission on user objects to a group
#====================================================================
function Grant-UserDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"CreateChild",$AccessControlTypeAllow,$GuidMap["user"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"DeleteChild",$AccessControlTypeAllow,$GuidMap["user"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,"Descendents",$GuidMap["user"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"ExtendedRight",$AccessControlTypeAllow,$ExtendedRightsMap["Reset Password"],"Descendents",$GuidMap["user"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permission on OU objects to a group
#====================================================================
function Grant-OUDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"CreateChild",$AccessControlTypeAllow,$GuidMap["organizationalUnit"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"DeleteChild",$AccessControlTypeAllow,$GuidMap["organizationalUnit"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,"Descendents",$GuidMap["organizationalUnit"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to DNS Operators
#====================================================================
function Grant-DNSOperatorsPermissionDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetDN
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AdminGroupSIDidentity = [System.Security.Principal.IdentityReference] $AdminGroupSID
    $adRightsGR = [System.DirectoryServices.ActiveDirectoryRights] "GenericRead"
    $adRightsGE = [System.DirectoryServices.ActiveDirectoryRights] "GenericExecute"
    $adRightsGW = [System.DirectoryServices.ActiveDirectoryRights] "GenericWrite"
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] "CreateChild"
    $adRightsDC = [System.DirectoryServices.ActiveDirectoryRights] "DeleteChild"
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "All"
    $Acl = Get-Acl "AD:\$TargetDN,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGR,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGE,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGW,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsCC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to DNS ReadOnly Group
#====================================================================
function Grant-DNSReadOnlyPermissionDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetDN
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AdminGroupSIDidentity = [System.Security.Principal.IdentityReference] $AdminGroupSID
    $adRightsGR = [System.DirectoryServices.ActiveDirectoryRights] "GenericRead"
    $adRightsGE = [System.DirectoryServices.ActiveDirectoryRights] "GenericExecute"
    $adRightsGW = [System.DirectoryServices.ActiveDirectoryRights] "GenericWrite"
    $adRightsGA = [System.DirectoryServices.ActiveDirectoryRights] "GenericAll"
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] "CreateChild"
    $adRightsDC = [System.DirectoryServices.ActiveDirectoryRights] "DeleteChild"
    $adRightsRC = [System.DirectoryServices.ActiveDirectoryRights] "ReadControl"
    $adRightsWO = [System.DirectoryServices.ActiveDirectoryRights] "WriteOwner"
    $adRightsWD = [System.DirectoryServices.ActiveDirectoryRights] "WriteDacl"
    $adRightsDT = [System.DirectoryServices.ActiveDirectoryRights] "DeleteTree"
    $adRightsDEL = [System.DirectoryServices.ActiveDirectoryRights] "Delete"
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $AccessControlTypeDeny = [System.Security.AccessControl.AccessControlType] "Deny"
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "All"
    $Acl = Get-Acl "AD:\$TargetDN,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGR,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGE,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGW,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsCC,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDC,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsWO,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsWD,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDT,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDEL,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.RemoveAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsRC,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to AD Sites, Subnets, and Transports admins
#====================================================================
function Grant-ADObjectPermissionDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetDN
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AdminGroupSIDidentity = [System.Security.Principal.IdentityReference] $AdminGroupSID
    $adRightsGA = [System.DirectoryServices.ActiveDirectoryRights] "GenericAll"
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] "CreateChild"
    $adRightsDC = [System.DirectoryServices.ActiveDirectoryRights] "DeleteChild"
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "All"
    $Acl = Get-Acl "AD:\$TargetDN,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGA,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsCC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to GPO admins
#====================================================================
function Grant-GPOPermissionDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AdminGroupSIDidentity = [System.Security.Principal.IdentityReference] $AdminGroupSID
    $adRightsRP = [System.DirectoryServices.ActiveDirectoryRights] "ReadProperty"
    $adRightsWP = [System.DirectoryServices.ActiveDirectoryRights] "WriteProperty"
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "All"
    $Acl = Get-Acl "AD:\$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsRP,$AccessControlTypeAllow,$GuidMap["gPLink"],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWP,$AccessControlTypeAllow,$GuidMap["gPLink"],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsRP,$AccessControlTypeAllow,$GuidMap["gPOptions"],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWP,$AccessControlTypeAllow,$GuidMap["gPOptions"],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"ExtendedRight",$AccessControlTypeAllow,$ExtendedRightsMap["Generate Resultant Set of Policy (Logging)"],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"ExtendedRight",$AccessControlTypeAllow,$ExtendedRightsMap["Generate Resultant Set of Policy (Planning)"],$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to create GPOs
#====================================================================
function Grant-GPOCreationDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetDN
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AdminGroupSIDidentity = [System.Security.Principal.IdentityReference] $AdminGroupSID
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] "CreateChild"
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $inheritanceTypeNone = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "None"
    $Acl = Get-Acl "AD:\$TargetDN,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsCC,$AccessControlTypeAllow,$inheritanceTypeNone))
    $Acl | Set-Acl
}
#====================================================================

Write-Log ("=" * 80)
Write-Log "Log file is '$LogFile'"
Write-Log ("=" * 80)
Write-Log "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log "DC being used is '$DCHostName'"
Write-Log "Script path is '$ScriptPath'"
Write-Log "$ScriptTitle"
Write-Log "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log ("=" * 80)
Write-Log ""
$requiredGroups = @('Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

#====================================================================
# Add additional UPN suffix
#====================================================================
Write-Log "Adding $EmailSuffix as an additional UPN suffix"
$forest = Get-ADForest
if ($forest.UPNSuffixes -notcontains $EmailSuffix) {
    Get-ADForest | Set-ADForest -UPNSuffixes @{add = $EmailSuffix} -Server $DCHostName
}
#====================================================================

#====================================================================
# Prevent standard users from creating computer accounts
#====================================================================
Write-Log "Preventing standard users from creating computer accounts"
Set-ADDomain (Get-ADDomain).distinguishedname -Replace @{"ms-ds-MachineAccountQuota"="0"} -Server $DCHostName
#====================================================================

#====================================================================
# Enable PAM feature to use temporal group membership
#====================================================================
Write-Log "Enabling PAM feature to use temporal group membership"
Enable-ADOptionalFeature "Privileged Access Management Feature" -Scope ForestOrConfigurationSet -Target $DNSSuffix -Confirm:$False -Server $DCHostName
#====================================================================

#====================================================================
# Enable AD Recycle Bin
#====================================================================
Write-Log "Enabling AD Recycle Bin"
Enable-ADOptionalFeature 'Recycle Bin Feature' -Scope ForestOrConfigurationSet -Target $DNSSuffix -Confirm:$False -Server $DCHostName
#====================================================================

#====================================================================
# Protect Domain Controllers OU
#====================================================================
Write-Log "Protecting Domain Controllers OU"
Set-ADOrganizationalUnit -Identity "OU=Domain Controllers,$EndPath" -ProtectedFromAccidentalDeletion $true -Server $DCHostName
#====================================================================

Write-Log "Creating user OUs & Groups"
#====================================================================
# Staff OU & group creation
#====================================================================
New-ADOU -OUName $StaffGroup -Path $EndPath -OUDescription "Top level OU for Staff User objects"
New-ADOU -OUName $GroupsOU -Path $EndPath -OUDescription "Top level OU for Group objects"
New-DomainGroup -GroupName $StaffGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Org-wide group for all staff users"
New-DomainGroup -GroupName $ITGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Departmental group holding all IT accounts"
New-DomainGroup -GroupName $O365LicenseGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Used to assign Office365 licenses"
New-ADOU -OUName "Administration" -Path $EndPath -OUDescription "Top level OU for IT Admin User & Group objects"
New-ADOU -OUName "Equipment_Mailbox_Access" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT Admin OU for Group objects that grant access to equipment mailboxes"
New-ADOU -OUName "Room_Mailbox_Access" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT Admin OU for Group objects that grant access to room mailboxes"
New-ADOU -OUName "Shared_Mailbox_Access" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT Admin OU for Group objects that grant access to shared mailboxes"
New-ADOU -OUName "Equipment_Mailbox_Accounts" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are equipment mailbox recipient types"
New-ADOU -OUName "Hi_Priv_Groups" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for Group objects that control Hi-Priv access"
New-ADOU -OUName "Hi_Priv_Accounts" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are Hi-Priv accounts"
New-ADOU -OUName "Local_Admin_Groups" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for Group objects that give local admin on individual devices"
New-ADOU -OUName "Room_Mailbox_Accounts" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are room mailbox recipient types"
New-ADOU -OUName "Shared_Mailbox_Accounts" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are shared mailbox recipient types"
New-ADOU -OUName "Service_Accounts" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are service accounts"
New-DomainGroup -GroupName $ITAdminGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Group holding all IT Admin accounts"
New-DomainGroup -GroupName "ADM_Task_AD_Administration_OU_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
New-DomainGroup -GroupName "ADM_Task_AD_Computer_OU_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
New-DomainGroup -GroupName "ADM_Task_AD_GPO_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create & link GPOs"
New-DomainGroup -GroupName "ADM_Task_AD_Group_OU_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
New-DomainGroup -GroupName "ADM_Task_AD_Site_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
New-DomainGroup -GroupName "ADM_Task_AD_Subnet_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Subnet objects"
New-DomainGroup -GroupName "ADM_Task_AD_Transport_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Transport objects"
New-DomainGroup -GroupName "ADM_Task_AD_User_OU_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
New-DomainGroup -GroupName "ADM_Task_Desktop_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members are added to Local Admin on all computers in the Desktop, Laptop, & VM OUs"
New-DomainGroup -GroupName "ADM_Task_DFS_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control NTFS permissions on all DFS folders & have access to DFS console"
New-DomainGroup -GroupName "ADM_Task_DHCP_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members are members of DHCP Administrators"
New-DomainGroup -GroupName "ADM_Task_DHCP_Users" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members are members of DHCP Users"
New-DomainGroup -GroupName $DNSOperatorsGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Edit access to DNS zones"
New-DomainGroup -GroupName $DNSReadOnlyGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have read only access to the DNS service"
New-DomainGroup -GroupName $HiPrivAccountAdminGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit accounts in the Hi_Priv_Accounts OU"
New-DomainGroup -GroupName $HiPrivGroupAdminGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit groups in the Hi_Priv_Groups OU"
New-DomainGroup -GroupName $InstallerGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have permission to join & move computer objects in $ParentOU & Sub OUs"
New-DomainGroup -GroupName $LocalAdminGroupAdminGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit groups in the Local_Admin_Groups OU"
New-DomainGroup -GroupName $UserPasswordDelegationGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can reset passwords of users in the $StaffGroup OU"
New-DomainGroup -GroupName "ADM_Task_Server_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members are added to Local admin on all computers in $ParentOU & Sub OUs via GPO, are indirect members of the Server Operators BuiltIn group"
New-DomainGroup -GroupName $SERAccessAdminGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can edit the membership of sh_, eq_, & ro_ groups"
New-DomainGroup -GroupName $SERAccountAdminGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit sh_, eq_, & ro_ accounts"
New-DomainGroup -GroupName $ServiceAccountAdminGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit accounts in the Service_Accounts OU"
New-DomainGroup -GroupName $StandardAccountAdminGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit accounts in the $StaffGroup OU"
New-DomainGroup -GroupName $StandardGroupAdminGroup -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit groups in the $GroupsOU OU"
New-DomainGroup -GroupName "ADM_Task_WDS_Deploy_Clients" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can use WDS to deploy images in the Clients folder"
New-DomainGroup -GroupName "ADM_Task_WDS_Deploy_Servers" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can use WDS to deploy images in the Servers folder"
New-DomainGroup -GroupName "ADM_Role_Tier1_Level_1_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Level 1 admins - desktop"
New-DomainGroup -GroupName "ADM_Role_Tier1_Level_2_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Level 2 admins - junior server"
New-DomainGroup -GroupName "ADM_Role_Tier1_Level_3_Admins" -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Level 3 admins - senior server"
Add-GroupMember -group $StaffGroup -Member $ITGroup
Add-GroupMember -group "Remote Desktop Users" -Member "ADM_Task_Server_Admins"
Add-GroupMember -group "Server Operators" -Member "ADM_Task_Server_Admins"
Add-GroupMember -group $ITAdminGroup -Member $SID500
Add-GroupMember -group "ADM_Task_AD_Administration_OU_Admins" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_Computer_OU_Admins" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_GPO_Admins" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_Group_OU_Admins" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_Site_Admins" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_Subnet_Admins" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_Transport_Admins" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_User_OU_Admins" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group "ADM_Task_Desktop_Admins" -Member "ADM_Role_Tier1_Level_1_Admins"
Add-GroupMember -group "ADM_Task_Desktop_Admins" -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group "ADM_Task_Desktop_Admins" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group "ADM_Task_DHCP_Admins" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group "ADM_Task_DHCP_Users" -Member "ADM_Role_Tier1_Level_1_Admins"
Add-GroupMember -group "ADM_Task_DHCP_Users" -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group "ADM_Task_DFS_Admins" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group $DNSOperatorsGroup -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group $DNSOperatorsGroup -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group $DNSReadOnlyGroup -Member "ADM_Role_Tier1_Level_1_Admins"
Add-GroupMember -group $HiPrivAccountAdminGroup -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group $HiPrivGroupAdminGroup -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group $InstallerGroup -Member "ADM_Role_Tier1_Level_1_Admins"
Add-GroupMember -group $InstallerGroup -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group $InstallerGroup -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group $LocalAdminGroupAdminGroup -Member "ADM_Role_Tier1_Level_1_Admins"
Add-GroupMember -group $LocalAdminGroupAdminGroup -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group $LocalAdminGroupAdminGroup -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group $UserPasswordDelegationGroup -Member "ADM_Role_Tier1_Level_1_Admins"
Add-GroupMember -group $UserPasswordDelegationGroup -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group $UserPasswordDelegationGroup -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group "ADM_Task_Server_Admins" -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group "ADM_Task_Server_Admins" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group $SERAccessAdminGroup -Member "ADM_Role_Tier1_Level_1_Admins"
Add-GroupMember -group $SERAccessAdminGroup -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group $SERAccessAdminGroup -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group $SERAccountAdminGroup -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group $SERAccountAdminGroup -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group $StandardAccountAdminGroup -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group $StandardAccountAdminGroup -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group $StandardGroupAdminGroup -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group $StandardGroupAdminGroup -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group $ServiceAccountAdminGroup -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group $ServiceAccountAdminGroup -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group "ADM_Task_WDS_Deploy_Clients" -Member "ADM_Role_Tier1_Level_1_Admins"
Add-GroupMember -group "ADM_Task_WDS_Deploy_Clients" -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group "ADM_Task_WDS_Deploy_Clients" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -group "ADM_Task_WDS_Deploy_Servers" -Member "ADM_Role_Tier1_Level_2_Admins"
Add-GroupMember -group "ADM_Task_WDS_Deploy_Servers" -Member "ADM_Role_Tier1_Level_3_Admins"
#====================================================================

Write-Log "Creating shares"
#====================================================================
# Create main Share
#====================================================================
$ShareName = $RootShare
if (!(TEST-PATH "\\$DNSSuffix\$ShareName")) {
    if (!(TEST-PATH "\\$ServerName\$ShareName")) {
        if (!(TEST-PATH "$Drive\$ShareName")) {
            New-Item "$Drive\$ShareName" -type directory
        } else {
            Write-Log "$Drive\$ShareName already exists" -ForegroundColor Green
        }
        New-SmbShare -Name $ShareName -Path "$Drive\$ShareName" -FullAccess "Administrators", "SYSTEM" -ChangeAccess "authenticated users"
        Write-Log "Pausing for 60 seconds after creating share $ShareName"
        Start-Sleep -s 60
    } else {
        Write-Log "\\$ServerName\$ShareName already exists" -ForegroundColor Green
    }
    New-DfsnRoot -TargetPath "\\$ServerName\$ShareName" -Type DomainV2 -Path "\\$DNSSuffix\$ShareName" -GrantAdminAccounts "ADM_Task_DFS_Admins"
    New-DfsReplicationGroup -GroupName $ShareName | New-DfsReplicatedFolder -FolderName $ShareName -DfsnPath "\\$DNSSuffix\$ShareName" | Add-DfsrMember -ComputerName $ServerName
    Set-DfsrMembership -GroupName $ShareName -FolderName $ShareName -ContentPath "$Drive\$ShareName" -ComputerName $ServerName -PrimaryMember $True -StagingPathQuotaInMB 16384 -Force
    Grant-DfsrDelegation -GroupName $ShareName -AccountName "ADM_Task_DFS_Admins" -Force
} else {
    Write-Log "\\$DNSSuffix\$ShareName already exists" -ForegroundColor Green
}
#====================================================================

#====================================================================
# Create Profiles Share
#====================================================================
$ShareName = "Profiles"
if (!(TEST-PATH "\\$DNSSuffix\$ShareName")) {
    if (!(TEST-PATH "\\$ServerName\$ShareName")) {
        if (!(TEST-PATH "$Drive\$ShareName")) {
            New-Item "$Drive\$ShareName" -type directory -force
            $Acl = Get-Acl "$Drive\$ShareName"
            $isProtected = $true
            $preserveInheritance = $false
            $Acl.SetAccessRuleProtection($isProtected, $preserveInheritance)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($StaffGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($StandardAccountAdminGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("ADM_Task_DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            Set-Acl "$Drive\$ShareName" $Acl
        } else {
            Write-Log "$Drive\$ShareName already exists" -ForegroundColor Green
        }
        New-SmbShare -Name $ShareName -Path "$Drive\$ShareName" -FullAccess "Administrators", "SYSTEM" -ChangeAccess "authenticated users"
        Write-Log "Pausing for 60 seconds after creating share $ShareName"
        Start-Sleep -s 60
    } else {
        Write-Log "\\$ServerName\$ShareName already exists" -ForegroundColor Green
    }
    New-DfsnRoot -TargetPath "\\$ServerName\$ShareName" -Type DomainV2 -Path "\\$DNSSuffix\$ShareName" -GrantAdminAccounts "ADM_Task_DFS_Admins"
    New-DfsReplicationGroup -GroupName $ShareName | New-DfsReplicatedFolder -FolderName $ShareName -DfsnPath "\\$DNSSuffix\$ShareName" | Add-DfsrMember -ComputerName $ServerName
    Set-DfsrMembership -GroupName $ShareName -FolderName $ShareName -ContentPath "$Drive\$ShareName" -ComputerName $ServerName -PrimaryMember $True -StagingPathQuotaInMB 16384 -Force
    Grant-DfsrDelegation -GroupName $ShareName -AccountName "ADM_Task_DFS_Admins" -Force
} else {
    Write-Log "\\$DNSSuffix\$ShareName already exists" -ForegroundColor Green
}
#====================================================================

#====================================================================
# Create IT Group Share
#====================================================================
$ShareName = $ITGroup
if (!(TEST-PATH "\\$DNSSuffix\$RootShare\$ShareName")) {
    New-Item "\\$DNSSuffix\$RootShare\$ShareName" -type directory
    $Acl = Get-Acl "\\$DNSSuffix\$RootShare\$ShareName"
    $isProtected = $true
    $preserveInheritance = $false
    $Acl.SetAccessRuleProtection($isProtected, $preserveInheritance)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($ShareName,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($ITAdminGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("ADM_Task_DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    Set-Acl "\\$DNSSuffix\$RootShare\$ShareName" $Acl
    Robocopy $UserScriptsLocation "\\$DNSSuffix\$RootShare\$ShareName\User_Scripts" /e
} else {
    Write-Log "\\$DNSSuffix\$RootShare\$ShareName already exists" -ForegroundColor Green
}
$ShareName = "Software"
if (!(TEST-PATH "\\$DNSSuffix\$RootShare\$ShareName")) {
    New-Item "\\$DNSSuffix\$RootShare\$ShareName" -type directory
    $Acl = Get-Acl "\\$DNSSuffix\$RootShare\$ShareName"
    $isProtected = $true
    $preserveInheritance = $false
    $Acl.SetAccessRuleProtection($isProtected, $preserveInheritance)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Authenticated Users","ReadAndExecute","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($ITGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($ITAdminGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("ADM_Task_DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    Set-Acl "\\$DNSSuffix\$RootShare\$ShareName" $Acl
} else {
    Write-Log "\\$DNSSuffix\$RootShare\$ShareName already exists" -ForegroundColor Green
}
#====================================================================

Write-Log "Creating computer OUs"
#====================================================================
# Create default computers OU
#====================================================================
New-ADOU -OUName $ParentOU -Path $EndPath -OUDescription "Top level OU for Computer objects"
#====================================================================

#====================================================================
# Create OU for servers
#====================================================================
New-ADOU -OUName "Servers" -Path $Location
#====================================================================

#====================================================================
# Create OU for desktops
#====================================================================
New-ADOU -OUName "Desktops" -Path $Location
#====================================================================

#====================================================================
# Create OU for laptops
#====================================================================
New-ADOU -OUName "Laptops" -Path $Location
#====================================================================

#====================================================================
# Create OU for VMs
#====================================================================
New-ADOU -OUName "VMs" -Path $Location
#====================================================================

#====================================================================
# Redirect default computer location & delegate permissions
#====================================================================
Write-Log "Creating Permission delegations"
$guidMap = @{}
$GuidMapParams = @{
    SearchBase = ($rootdse.SchemaNamingContext)
    LDAPFilter = "(schemaidguid=*)"
    Properties = ("lDAPDisplayName", "schemaIDGUID")
}
Get-ADObject @GuidMapParams -Server $DCHostName | ForEach-Object { $guidMap[$_.lDAPDisplayName] = [System.GUID]$_.schemaIDGUID }
$ExtendedRightsMap = @{}
$ExtendedMapParams = @{
    SearchBase = ($rootdse.ConfigurationNamingContext)
    LDAPFilter = "(&(objectclass=controlAccessRight)(rightsguid=*))"
    Properties = ("displayName", "rightsGuid")
}
Get-ADObject @ExtendedMapParams -Server $DCHostName | ForEach-Object { $ExtendedRightsMap[$_.displayName] = [System.GUID]$_.rightsGuid }
redircmp $Location
Grant-ComputerJoinDelegation -AdminGroupName $InstallerGroup -TargetOU $ParentOU
Grant-OUDelegation -AdminGroupName "ADM_Task_AD_Computer_OU_Admins" -TargetOU $ParentOU
Grant-PasswordResetDelegation -AdminGroupName $UserPasswordDelegationGroup -TargetOU $StaffGroup
Grant-GroupMembershipEditDelegation -AdminGroupName $SERAccessAdminGroup -TargetOU $EquipmentGroupsOU
Grant-GroupMembershipEditDelegation -AdminGroupName $SERAccessAdminGroup -TargetOU $RoomGroupsOU
Grant-GroupMembershipEditDelegation -AdminGroupName $SERAccessAdminGroup -TargetOU $SharedGroupsOU
Grant-UserDelegation -AdminGroupName $SERAccountAdminGroup -TargetOU $EquipmentAccountsOU
Grant-GroupDelegation -AdminGroupName $SERAccountAdminGroup -TargetOU $EquipmentGroupsOU
Grant-UserDelegation -AdminGroupName $SERAccountAdminGroup -TargetOU $RoomAccountsOU
Grant-GroupDelegation -AdminGroupName $SERAccountAdminGroup -TargetOU $RoomGroupsOU
Grant-UserDelegation -AdminGroupName $SERAccountAdminGroup -TargetOU $SharedAccountsOU
Grant-GroupDelegation -AdminGroupName $SERAccountAdminGroup -TargetOU $SharedGroupsOU
Grant-UserDelegation -AdminGroupName $HiPrivAccountAdminGroup -TargetOU "Hi_Priv_Accounts,OU=$AdministrationOU"
Grant-GroupDelegation -AdminGroupName $HiPrivGroupAdminGroup -TargetOU "Hi_Priv_Groups,OU=$AdministrationOU"
Grant-UserDelegation -AdminGroupName $StandardAccountAdminGroup -TargetOU $StaffGroup
Grant-GroupDelegation -AdminGroupName $StandardGroupAdminGroup -TargetOU $GroupsOU
Grant-UserDelegation -AdminGroupName $ServiceAccountAdminGroup -TargetOU "Service_Accounts,OU=$AdministrationOU"
Grant-GroupDelegation -AdminGroupName $LocalAdminGroupAdminGroup -TargetOU "Local_Admin_Groups,OU=$AdministrationOU"
Grant-OUDelegation -AdminGroupName "ADM_Task_AD_Administration_OU_Admins" -TargetOU $AdministrationOU
Grant-OUDelegation -AdminGroupName "ADM_Task_AD_Group_OU_Admins" -TargetOU $GroupsOU
Grant-OUDelegation -AdminGroupName "ADM_Task_AD_User_OU_Admins" -TargetOU $StaffGroup
Grant-ADObjectPermissionDelegation -AdminGroupName "ADM_Task_AD_Site_Admins" -TargetDN "CN=Sites,CN=Configuration"
Grant-ADObjectPermissionDelegation -AdminGroupName "ADM_Task_AD_Subnet_Admins" -TargetDN "CN=Subnets,CN=Sites,CN=Configuration"
Grant-ADObjectPermissionDelegation -AdminGroupName "ADM_Task_AD_Transport_Admins" -TargetDN "CN=Inter-Site Transports,CN=Sites,CN=Configuration"
Grant-GPOPermissionDelegation -AdminGroupName "ADM_Task_AD_GPO_Admins"
Grant-GPOCreationDelegation -AdminGroupName "ADM_Task_AD_GPO_Admins" -TargetDN "CN=Policies,CN=System"
Grant-DNSOperatorsPermissionDelegation -AdminGroupName $DNSOperatorsGroup -TargetDN "CN=MicrosoftDNS,CN=System"
Grant-DNSOperatorsPermissionDelegation -AdminGroupName $DNSOperatorsGroup -TargetDN "CN=MicrosoftDNS,DC=DomainDnsZones"
Grant-DNSReadOnlyPermissionDelegation -AdminGroupName $DNSReadOnlyGroup -TargetDN "CN=MicrosoftDNS,CN=System"
Grant-DNSReadOnlyPermissionDelegation -AdminGroupName $DNSReadOnlyGroup -TargetDN "CN=MicrosoftDNS,DC=DomainDnsZones"
$DNSZones = Get-ADObject -Filter * -SearchBase "CN=MicrosoftDNS,DC=DomainDnsZones,$EndPath" -SearchScope 1
foreach ($DNSZone in $DNSZones) {
    $DNSZoneName = $DNSZone.Name
    Grant-DNSReadOnlyPermissionDelegation -AdminGroupName $DNSReadOnlyGroup -TargetDN "DC=$DNSZoneName,CN=MicrosoftDNS,DC=DomainDnsZones"
}
#====================================================================

#====================================================================
# Set up LAPS
#====================================================================
Write-Log "Setting up LAPS"
Update-LapsADSchema -Confirm:$False
Set-LapsADComputerSelfPermission -Identity $Location
Set-LapsADReadPasswordPermission -Identity "OU=Desktops,$Location" -AllowedPrincipals "$Domain\ADM_Task_Desktop_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=Desktops,$Location" -AllowedPrincipals "$Domain\ADM_Task_Desktop_Admins"
Set-LapsADReadPasswordPermission -Identity "OU=Laptops,$Location" -AllowedPrincipals "$Domain\ADM_Task_Desktop_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=Laptops,$Location" -AllowedPrincipals "$Domain\ADM_Task_Desktop_Admins"
Set-LapsADReadPasswordPermission -Identity "OU=Servers,$Location" -AllowedPrincipals "$Domain\ADM_Task_Server_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=Servers,$Location" -AllowedPrincipals "$Domain\ADM_Task_Server_Admins"
Set-LapsADReadPasswordPermission -Identity "OU=VMs,$Location" -AllowedPrincipals "$Domain\ADM_Task_Desktop_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=VMs,$Location" -AllowedPrincipals "$Domain\ADM_Task_Desktop_Admins"
#====================================================================

#====================================================================
# Secure built-in admin account
#====================================================================
Write-Log "Securing built-in admin account"
Move-ADObject -Identity $(Get-ADUser -Identity $SID500).DistinguishedName -TargetPath "OU=Hi_Priv_Accounts,OU=$AdministrationOU,$EndPath" -Server $DCHostName
Set-ADAccountControl -Identity $SID500 -AccountNotDelegated $True -Server $DCHostName
Remove-ADGroupMember -Identity "Schema Admins" -Members $SID500 -Server $DCHostName -Confirm:$False
#====================================================================

#====================================================================
# Import GPOs
#====================================================================
Write-Log "Creating & linking GPOs"
$GPOName = "Default Domain Policy"
try {
    Set-GPLink -name $GPOName -target $EndPath -enforced yes -ErrorAction Stop -Server $DCHostName
    Write-Log "Enforced $GPOName on $GPOTarget"
} catch {
    $ex = $_.Exception
    throw
}
$GPOName = "Default Domain Controllers Policy"
try {
    Set-GPLink -name $GPOName -target "ou=Domain Controllers,$EndPath" -enforced yes -ErrorAction Stop -Server $DCHostName
    Write-Log "Enforced $GPOName on $GPOTarget"
} catch {
    $ex = $_.Exception
    throw
}
$ImportedGPOs += Import-GPO -BackupGpoName "Logon Policy" -TargetName "Logon Policy" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "Logon Policy" -GPOTarget $EndPath
$ImportedGPOs += Import-GPO -BackupGpoName "TLS" -TargetName "TLS" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "TLS" -GPOTarget $EndPath
$ImportedGPOs += Import-GPO -BackupGpoName "CWDIllegalInDllSearch" -TargetName "CWDIllegalInDllSearch" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "CWDIllegalInDllSearch" -GPOTarget $EndPath
$ImportedGPOs += Import-GPO -BackupGpoName "DisableNullSessionEnumeration" -TargetName "DisableNullSessionEnumeration" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "DisableNullSessionEnumeration" -GPOTarget $EndPath
$ImportedGPOs += Import-GPO -BackupGpoName "EnableSMBSigning" -TargetName "EnableSMBSigning" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "EnableSMBSigning" -GPOTarget $EndPath
$ImportedGPOs += Import-GPO -BackupGpoName "EnforceNLAandTLSforRDP" -TargetName "EnforceNLAandTLSforRDP" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "EnforceNLAandTLSforRDP" -GPOTarget $EndPath
$ImportedGPOs += Import-GPO -BackupGpoName "GroupPolicyHardening_MS150-011" -TargetName "GroupPolicyHardening_MS150-011" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "GroupPolicyHardening_MS150-011" -GPOTarget $EndPath
$ImportedGPOs += Import-GPO -BackupGpoName "Kerberos_Armouring" -TargetName "Kerberos_Armouring" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "Kerberos_Armouring" -GPOTarget $EndPath
$ImportedGPOs += Import-GPO -BackupGpoName "LSAProtection" -TargetName "LSAProtection" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "LSAProtection" -GPOTarget $EndPath
$ImportedGPOs += Import-GPO -BackupGpoName "NTLMv2" -TargetName "NTLMv2" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "NTLMv2" -GPOTarget $EndPath
$ImportedGPOs += Import-GPO -BackupGpoName "LDAP_signing_requirements" -TargetName "LDAP_signing_requirements" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "LDAP_signing_requirements" -GPOTarget $EndPath
$ImportedGPOs += Import-GPO -BackupGpoName "DC_AppLocker_Disable_Browsers" -TargetName "DC_AppLocker_Disable_Browsers" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "DC_AppLocker_Disable_Browsers" -GPOTarget "ou=Domain Controllers,$EndPath"
$ImportedGPOs += Import-GPO -BackupGpoName "DC_Auditing" -TargetName "DC_Auditing" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "DC_Auditing" -GPOTarget "ou=Domain Controllers,$EndPath"
$ImportedGPOs += Import-GPO -BackupGpoName "DC_Disable_Print_Spooler" -TargetName "DC_Disable_Print_Spooler" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "DC_Disable_Print_Spooler" -GPOTarget "ou=Domain Controllers,$EndPath"
$ImportedGPOs += Import-GPO -BackupGpoName "DC_LDAP_signing_requirements" -TargetName "DC_LDAP_signing_requirements" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "DC_LDAP_signing_requirements" -GPOTarget "ou=Domain Controllers,$EndPath"
$ImportedGPOs += Import-GPO -BackupGpoName "ADM_Task_Server_Admins as members of Local admins" -TargetName "ADM_Task_Server_Admins as members of Local admins" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "ADM_Task_Server_Admins as members of Local admins" -GPOTarget $Location
$ImportedGPOs += Import-GPO -BackupGpoName "LAPS" -TargetName "LAPS" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "LAPS" -GPOTarget $Location
$ImportedGPOs += Import-GPO -BackupGpoName "ADM_Task_Desktop_Admins as members of Local admins" -TargetName "ADM_Task_Desktop_Admins as members of Local admins" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "ADM_Task_Desktop_Admins as members of Local admins" -GPOTarget "OU=Desktops,$Location"
Add-GPOLink -GPOName "ADM_Task_Desktop_Admins as members of Local admins" -GPOTarget "OU=Laptops,$Location"
Add-GPOLink -GPOName "ADM_Task_Desktop_Admins as members of Local admins" -GPOTarget "OU=VMs,$Location"
$ImportedGPOs += Import-GPO -BackupGpoName "IT Desktop Prefs" -TargetName "IT Desktop Prefs" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "IT Desktop Prefs" -GPOTarget "OU=$StaffGroup,$EndPath"
Add-GPOLink -GPOName "IT Desktop Prefs" -GPOTarget "OU=Hi_Priv_Accounts,OU=$AdministrationOU,$EndPath"
Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoRead -TargetName "Authenticated Users" -TargetType Group -Server $DCHostName -Replace
Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoApply -TargetName $ITGroup -TargetType Group -Server $DCHostName
Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoApply -TargetName $ITAdminGroup -TargetType Group -Server $DCHostName
$ImportedGPOs += Import-GPO -BackupGpoName "CM visual help" -TargetName "CM visual help" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
Add-GPOLink -GPOName "CM visual help" -GPOTarget "OU=$StaffGroup,$EndPath"
Add-GPOLink -GPOName "CM visual help" -GPOTarget "OU=Hi_Priv_Accounts,OU=$AdministrationOU,$EndPath"
Set-GPPermission -Name "CM visual help" -PermissionLevel GpoRead -TargetName "Authenticated Users" -TargetType Group -Server $DCHostName -Replace
Set-GPPermission -Name "CM visual help" -PermissionLevel GpoApply -TargetName $SID500 -TargetType User -Server $DCHostName
foreach ($GPOName in $ImportedGPOs) {
    Set-GPPermission -Name $GPOName.DisplayName -PermissionLevel GpoEditDeleteModifySecurity -TargetName "ADM_Task_AD_GPO_Admins" -TargetType Group -Server $DCHostName
}
#====================================================================
