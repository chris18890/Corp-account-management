#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
    , [Parameter(Mandatory)][string]$UserName
    , [string]$FirstName,[string]$LastName,[string]$Description
    , [string]$Dept,[string]$Company,[string]$Manager
    , [string]$LogFile,[string]$DCHostName,[int]$PasswordLength
)

Add-Type -Assembly System.Web

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
    # Effects:          User will be added to the group
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
# Test password against password policy
#====================================================================
function Test-Password {
    param([string]$Password)
    #================================================================
    # Purpose:          Test password against password policy
    # Assumptions:      Password has been generated with enough characters for required groups
    # Effects:          Password should be valid
    # Inputs:           $Password
    # Calls:            Write-Log function
    # Returns:
    # Notes:            There are 4 requirements in the current policy, but this could change in future
    #================================================================
    $TestsPassed = 0
    if ($Password.length -ge ($PasswordLength)) {$TestsPassed ++} # Must be >= 15 characters in length
    if ($Password -cmatch "[a-z]") {$TestsPassed ++} # Must contain a lowercase letter
    if ($Password -cmatch "[A-Z]") {$TestsPassed ++} # Must contain an uppercase letter
    if ($Password -cmatch "[0-9]") {$TestsPassed ++} # Must contain a digit
    #if (-Not($Password -notmatch "[a-zA-Z0-9]")) {$TestsPassed ++} # Must contain a special character
    if ($TestsPassed -ge 4) {
        Write-Log "Password validated"
        Write-Log ""
    } else {
        Write-Log ("-" * 80) -ForegroundColor Red
        Write-Log "ERROR: Password does not comply with the password policy, script terminating" -ForegroundColor Red
        Write-Log ("-" * 80) -ForegroundColor Red
        return
    }
}
#====================================================================

#====================================================================
# Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
# ADConnect & Exchange settings
if (!$DCHostName) {
    $DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
}
# Get containing folder for script to locate supporting files
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
# Set variables
$ScriptTitle = "$Domain Domain Admin User Creation Script"
$OU = "Administration"
$SubOU = "Hi_Priv_Accounts"
$ITAdminGroup = "IT_Admin"
$OUPath = "OU=$SubOU,OU=$OU,$EndPath"
if (!$PasswordLength) {
    $PasswordLength = 20
}
if (!$Company) {
    $Company = $Domain
}
if (!$Dept) {
    $Dept = "IT"
}
# File locations
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Log "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
#====================================================================

if (!$LogFile) {
    $LogFileName = $Domain + "_new_Domain_Admin_user_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
    Write-Log ("=" * 80)
    Write-Log "Log file is '$LogFile'"
    Write-Log "Processing commenced, running as user '$Domain\$env:USERNAME'"
    Write-Log "Script path is '$ScriptPath'"
    Write-Log "$ScriptTitle"
    Write-Log "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
    Write-Log ("=" * 80)
    Write-Log ""
}
#====================================================================

$requiredGroups = @('ADM_Task_HiPriv_Account_Admins', 'ADM_Task_HiPriv_Group_Admins', 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

#====================================================================
# Set user info
#====================================================================
$PrivGroup = "Domain Admins"
if (!$FirstName) {
    $FirstName = READ-HOST 'Enter First Name - '
    $FirstName = $FirstName.Trim() -replace '[?@\\+]', [String]::Empty # Strip out illegal characters from First name for Office 365 compliance. Note that \ is escaped to \\
}
if (!$LastName) {
    $LastName = READ-HOST 'Enter Last Name - '
    $LastName = $LastName.Trim() -replace '[?@\\+]', [String]::Empty # Strip out illegal characters from LastName for Office 365 compliance. Note that \ is escaped to \\
}
if (!$UserName) {
    $UserName = READ-HOST 'Enter Username - '
    $UserName = $UserName.Trim() -replace '[^A-Za-z0-9.-]', [String]::Empty # Strip out illegal characters from User ID
    $UserNameDomainAdmin = "da." + $UserName
} else {
    $UserNameDomainAdmin = "da." + $UserName
}
if ($UserNameDomainAdmin.Length -gt 20) {
    $UserNameDomainAdmin = $UserNameDomainAdmin.Substring(0,20)
}
if (!$Manager) {
    $Manager = READ-HOST 'Enter manager username - '
    $Manager = $Manager.Trim() -replace '[^A-Za-z0-9.-]', [String]::Empty # Strip out illegal characters from User ID
}
$DisplayName = "$LastName, $FirstName (Domain Admin)"
$EmailAddress = "$UserNameDomainAdmin@$EmailSuffix"
$HomeDrive = "H:"
$HomeDir = "\\$DNSSuffix\Profiles\$UserNameDomainAdmin"
$UserPrincipalName = "$UserNameDomainAdmin@$EmailSuffix"
$ManagerDN = (Get-ADUser $Manager -Server $DCHostName).DistinguishedName

#====================================================================
$ExistingUser = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$UserNameDomainAdmin))" -Server $DCHostName
if ($ExistingUser) {
    Write-Log ("-" * 80) -ForegroundColor Red
    Write-Log "ERROR: User '$($UserNameDomainAdmin)' already exists in the $Domain directory. The user object`n is '$($ExistingUser.DistinguishedName)'" -ForegroundColor Red
    Write-Log "Error: Processing '$LastName, $FirstName' ($UserNameDomainAdmin) aborted" -ForegroundColor Red
    Write-Log ("-" * 80) -ForegroundColor Red
    Write-Log ""
    return # Skip this user
} else {
    try {
        Write-Log ("=" * 80)
        Write-Log "Processing '$DisplayName' ($UserNameDomainAdmin)..."
        Write-Log ("=" * 80)
        # Generate random password
        $UserPassword = [Web.Security.Membership]::GeneratePassword($PasswordLength,4)
        # Test password against password policy
        Test-Password -Password $UserPassword
        $Params = @{
            Name                    = $UserNameDomainAdmin
            AccountPassword         = ConvertTo-SecureString -AsPlainText $UserPassword -Force
            ChangePasswordAtLogon   = $true
            Company                 = $Company
            Department              = $Dept
            Description             = $Description
            DisplayName             = $DisplayName
            EmailAddress            = $EmailAddress
            Enabled                 = $true
            GivenName               = $FirstName
            HomeDirectory           = $HomeDir
            HomeDrive               = $HomeDrive
            ProfilePath             = $HomeDir
            Path                    = $OUPath
            Manager                 = $ManagerDN
            SamAccountName          = $UserNameDomainAdmin
            SurName                 = $LastName
            UserPrincipalName       = $UserPrincipalName
        }
        Write-Log "Creating $UserNameDomainAdmin"
        New-ADUser -Type "user" -Server $DCHostName @Params
        $UserPassword = $null
        Set-ADAccountControl -AccountNotDelegated $true -AllowReversiblePasswordEncryption $false -CannotChangePassword $false -DoesNotRequirePreAuth $false -Identity "CN=$UserNameDomainAdmin,$OUPath" -PasswordNeverExpires $false -UseDESKeyOnly $false -Server $DCHostName
        Set-ADObject -Identity "CN=$UserNameDomainAdmin,$OUPath" -protectedFromAccidentalDeletion $True -Server $DCHostName
        Add-GroupMember -Group $ITAdminGroup -Member $UserNameDomainAdmin
        Add-GroupMember -Group $PrivGroup -Member $UserNameDomainAdmin
        Add-GroupMember -Group "Protected Users" -Member $UserNameDomainAdmin
        $null = New-Item -Path $HomeDir -ItemType Directory -Force -ErrorAction SilentlyContinue
        try {
            $Acl = Get-Acl $HomeDir
            $Ar = New-Object System.Security.AccessControl.FileSystemAccessRule("$Domain\$UserNameDomainAdmin","Modify","ContainerInherit,ObjectInherit","None","Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            Set-Acl $HomeDir $Acl
            Write-Log "Created home directory $HomeDir"
        } catch {
            Write-Log "WARNING: Could not set ACL on $HomeDir - $($_.Exception.Message)" -ForegroundColor Yellow
        }
        Write-Log ("=" * 80)
        Write-Log "Processing for '$DisplayName' ($UserNameDomainAdmin) complete"
        Write-Log ("=" * 80)
        Write-Log ""
    } catch {
        $e = $_.Exception
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log ("-" * 80) -ForegroundColor Red
        Write-Log "ERROR: Failed during processing of '$DisplayName' ($UserNameDomainAdmin) - Line $Line" -ForegroundColor Red
        Write-Log "$e"
        Write-Log ("-" * 80) -ForegroundColor Red
        Write-Log ("=" * 80)
        Write-Log ""
    }
}
#====================================================================
