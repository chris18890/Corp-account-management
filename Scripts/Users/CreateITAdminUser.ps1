#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
    , [Parameter(Mandatory)][string]$UserName
    , [string]$FirstName,[string]$LastName,[string]$Description
    , [string]$Dept,[string]$Company,[string]$Manager
    , [string]$LogFile,[string]$DCHostName,[int]$PasswordLength
    , [ValidateSet(1,2,3)][int]$PrivLevel
)

Set-StrictMode -Version Latest
Add-Type -Assembly System.Web

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

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
$ScriptTitle = "$Domain HiPriv IT User Creation Script"
$OU = "Administration"
$SubOU = "Hi_Priv_Accounts"
$ITAdminGroup = "IT_Admin"
$OUPath = "OU=$SubOU,OU=$OU,$EndPath"
if (!$PrivLevel) {
    $PrivLevel = READ-HOST 'Enter a Privilege Level for the new account (1-3) - '
}
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
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
#====================================================================

if (!$LogFile) {
    $LogFileName = $Domain + "_new_Hi-Priv_user_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
    Write-Log -LogFile $LogFile -LogString ("=" * 80)
    Write-Log -LogFile $LogFile -LogString "Log file is '$LogFile'"
    Write-Log -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
    Write-Log -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
    Write-Log -LogFile $LogFile -LogString "$ScriptTitle"
    Write-Log -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
    Write-Log -LogFile $LogFile -LogString ("=" * 80)
    Write-Log -LogFile $LogFile -LogString " "
}
#====================================================================

$requiredGroups = @('ADM_Task_HiPriv_Account_Admins', 'ADM_Task_HiPriv_Group_Admins', 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

#====================================================================
# Set user info
#====================================================================
$PrivGroup = "ADM_Role_Tier1_Level_" + $PrivLevel + "_Admins"
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
    $UserNameAdmin = "admin." + $UserName
} else {
    $UserNameAdmin = "admin." + $UserName
}
if ($UserNameAdmin.Length -gt 20) {
    $UserNameAdmin = $UserNameAdmin.Substring(0,20)
}
if (!$Manager) {
    $Manager = READ-HOST 'Enter manager username - '
    $Manager = $Manager.Trim() -replace '[^A-Za-z0-9.-]', [String]::Empty # Strip out illegal characters from User ID
}
$DisplayName = "$LastName, $FirstName (Admin)"
$EmailAddress = "$UserNameAdmin@$EmailSuffix"
$HomeDrive = "H:"
$HomeDir = "\\$DNSSuffix\Profiles\$UserNameAdmin"
$UserPrincipalName = "$UserNameAdmin@$EmailSuffix"
$CheckManager = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$Manager))" -Server $DCHostName
if ($CheckManager) {
    $ManagerDN = (Get-ADUser $Manager -Server $DCHostName).DistinguishedName
} else {
    Write-Log -LogFile $LogFile -LogString "ERROR: Manager '$($Manager)' not found in the $Domain directory" -ForegroundColor Red
    $ManagerDN = $null
}

#====================================================================
$ExistingUser = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$UserNameAdmin))" -Server $DCHostName
if ($ExistingUser) {
    Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
    Write-Log -LogFile $LogFile -LogString "ERROR: User '$($UserNameAdmin)' already exists in the $Domain directory. The user object`n is '$($ExistingUser.DistinguishedName)'" -ForegroundColor Red
    Write-Log -LogFile $LogFile -LogString "Error: Processing '$LastName, $FirstName' ($UserNameAdmin) aborted" -ForegroundColor Red
    Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
    Write-Log -LogFile $LogFile -LogString " "
    return # Skip this user
} else {
    try {
        Write-Log -LogFile $LogFile -LogString ("=" * 80)
        Write-Log -LogFile $LogFile -LogString "Processing '$DisplayName' ($UserNameAdmin)..."
        Write-Log -LogFile $LogFile -LogString ("=" * 80)
        # Generate random password
        $UserPassword = [Web.Security.Membership]::GeneratePassword($PasswordLength,4)
        # Test password against password policy
        Test-Password -LogFile $LogFile -Password $UserPassword -PasswordLength $PasswordLength
        $Params = @{
            Name                    = $UserNameAdmin
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
            SamAccountName          = $UserNameAdmin
            SurName                 = $LastName
            UserPrincipalName       = $UserPrincipalName
        }
        Write-Log -LogFile $LogFile -LogString "Creating $UserNameAdmin"
        New-ADUser -Type "user" -Server $DCHostName @Params
        $UserPassword = $null
        Set-ADAccountControl -AccountNotDelegated $true -AllowReversiblePasswordEncryption $false -CannotChangePassword $false -DoesNotRequirePreAuth $false -Identity "CN=$UserNameAdmin,$OUPath" -PasswordNeverExpires $false -UseDESKeyOnly $false -Server $DCHostName
        Set-ADObject -Identity "CN=$UserNameAdmin,$OUPath" -protectedFromAccidentalDeletion $True -Server $DCHostName
        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $ITAdminGroup -Member $UserNameAdmin
        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $PrivGroup -Member $UserNameAdmin
        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Protected Users" -Member $UserNameAdmin
        $null = New-Item -Path $HomeDir -ItemType Directory -Force -ErrorAction SilentlyContinue
        try {
            $Acl = Get-Acl $HomeDir
            $Ar = New-Object System.Security.AccessControl.FileSystemAccessRule("$Domain\$UserNameAdmin","Modify","ContainerInherit,ObjectInherit","None","Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            Set-Acl $HomeDir $Acl
            Write-Log -LogFile $LogFile -LogString "Created home directory $HomeDir"
        } catch {
            Write-Log -LogFile $LogFile -LogString "WARNING: Could not set ACL on $HomeDir - $($_.Exception.Message)" -ForegroundColor Yellow
        }
        if ($PrivLevel -ge 2) {
            Write-Log -LogFile $LogFile -LogString "Creating Domain Admin account for $UserName"
            Write-Log -LogFile $LogFile -LogString " "
            & $PSScriptRoot\CreateITDomainAdminUser.ps1 -FirstName $FirstName -LastName $LastName -UserName $UserName -EmailSuffix $EmailSuffix -PrivLevel $PrivLevel -Description $Description -Dept $Dept -Company $Company -LogFile $LogFile -DCHostName $DCHostName -Manager $Manager -PasswordLength $PasswordLength
        }
        Write-Log -LogFile $LogFile -LogString ("=" * 80)
        Write-Log -LogFile $LogFile -LogString "Processing for '$DisplayName' ($UserNameAdmin) complete"
        Write-Log -LogFile $LogFile -LogString ("=" * 80)
        Write-Log -LogFile $LogFile -LogString " "
    } catch {
        $e = $_.Exception
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
        Write-Log -LogFile $LogFile -LogString "ERROR: Failed during processing of '$DisplayName' ($UserNameAdmin) - Line $Line" -ForegroundColor Red
        Write-Log -LogFile $LogFile -LogString "$e"
        Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
        Write-Log -LogFile $LogFile -LogString ("=" * 80)
        Write-Log -LogFile $LogFile -LogString " "
    }
}
#====================================================================
