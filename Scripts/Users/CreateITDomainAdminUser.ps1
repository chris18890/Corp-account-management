#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
    , [Parameter(Mandatory)][string]$UserName
    , [string]$FirstName,[string]$LastName,[string]$Description
    , [string]$Dept,[string]$Company,[string]$Manager
    , [string]$LogFile,[int]$PasswordLength,[string]$DCHostName
    , [ValidateSet(2,3)][int]$PrivLevel
)

Set-StrictMode -Version Latest

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

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
$ScriptPath = $PSScriptRoot
# Set variables
$ScriptTitle = "$Domain Domain Admin User Creation Script"
$OU = $Env.OUs.Administration
$SubOU = $Env.OUs.HiPrivAccounts
$ITAdminGroup = $Env.Groups.ITAdmin
$OUPath = "OU=$SubOU,OU=$OU,$EndPath"
if (!$PasswordLength) {
    $PasswordLength = $Env.Security.PasswordLength
}
if (!$Company) {
    $Company = $Domain
}
if (!$Dept) {
    $Dept = $Env.Groups.IT
}
# File locations
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_new_Domain_Admin_user_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-LogFile -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-LogFile -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-LogFile -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-LogFile -LogFile $LogFile -LogString "$ScriptTitle"
Write-LogFile -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "
#====================================================================

$requiredGroups = @("$($Env.Groups.TaskPrefix)HiPriv_Account_Admins", "$($Env.Groups.TaskPrefix)HiPriv_Group_Admins", 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

#====================================================================
# Set user info
#====================================================================
$PrivGroup = "$($Env.Groups.RolePrefix)Tier0_Level_" + $PrivLevel + "_Admins"
if (!$FirstName) {
    $FirstName = ConvertTo-SafeName (READ-HOST 'Enter First Name - ')
}
if (!$LastName) {
    $LastName = ConvertTo-SafeName (READ-HOST 'Enter Last Name - ')
}
if (!$UserName) {
    $UserName = READ-HOST 'Enter Username - '
}
$UserNameDomainAdmin = ConvertTo-SafeSamAccountName $UserName -Prefix 'da.'
if (!$Manager) {
    $Manager = ConvertTo-SafeSamAccountName (READ-HOST 'Enter manager username - ')
}
$DisplayName = "$LastName, $FirstName (Domain Admin)"
$EmailAddress = "$UserNameDomainAdmin@$EmailSuffix"
$HomeDrive = "H:"
$HomeDir = "\\$DNSSuffix\Profiles\$UserNameDomainAdmin"
$UserPrincipalName = "$UserNameDomainAdmin@$EmailSuffix"
$CheckManager = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$Manager))" -Server $DCHostName
$ManagerDN = $null
if ($CheckManager) {
    $ManagerDN = (Get-ADUser $Manager -Server $DCHostName).DistinguishedName
} else {
    Write-LogFile -LogFile $LogFile -LogString "WARNING: Manager '$($Manager)' not found in the $Domain directory - account created without manager" -ForegroundColor Yellow
}

#====================================================================
$ExistingUser = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$UserNameDomainAdmin))" -Server $DCHostName
if ($ExistingUser) {
    Write-LogFile -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
    Write-LogFile -LogFile $LogFile -LogString "ERROR: User '$($UserNameDomainAdmin)' already exists in the $Domain directory. The user object`n is '$($ExistingUser.DistinguishedName)'" -ForegroundColor Red
    Write-LogFile -LogFile $LogFile -LogString "Error: Processing '$LastName, $FirstName' ($UserNameDomainAdmin) aborted" -ForegroundColor Red
    Write-LogFile -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
    Write-LogFile -LogFile $LogFile -LogString " "
    return # Skip this user
} else {
    try {
        Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
        Write-LogFile -LogFile $LogFile -LogString "Processing '$DisplayName' ($UserNameDomainAdmin)..."
        Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
        # Generate random password
        $UserPassword = New-Password -Length $PasswordLength
        # Test password against password policy
        Test-Password -LogFile $LogFile -Password $UserPassword -PasswordLength $PasswordLength
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
            SamAccountName          = $UserNameDomainAdmin
            SurName                 = $LastName
            UserPrincipalName       = $UserPrincipalName
        }
        if ($ManagerDN) {$Params.Manager = $ManagerDN}
        Write-LogFile -LogFile $LogFile -LogString "Creating $UserNameDomainAdmin"
        New-ADUser -Type "user" -Server $DCHostName @Params
        $UserPassword = $null
        Set-ADAccountControl -AccountNotDelegated $true -AllowReversiblePasswordEncryption $false -CannotChangePassword $false -DoesNotRequirePreAuth $false -Identity "CN=$UserNameDomainAdmin,$OUPath" -PasswordNeverExpires $false -UseDESKeyOnly $false -Server $DCHostName
        Set-ADObject -Identity "CN=$UserNameDomainAdmin,$OUPath" -protectedFromAccidentalDeletion $True -Server $DCHostName
        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $ITAdminGroup -Member $UserNameDomainAdmin
        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $PrivGroup -Member $UserNameDomainAdmin
        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Protected Users" -Member $UserNameDomainAdmin
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
            Write-LogFile -LogFile $LogFile -LogString "Created home directory $HomeDir"
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "WARNING: Could not set ACL on $HomeDir - $($_.Exception.Message)" -ForegroundColor Yellow
        }
        Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
        Write-LogFile -LogFile $LogFile -LogString "Processing for '$DisplayName' ($UserNameDomainAdmin) complete"
        Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
        Write-LogFile -LogFile $LogFile -LogString " "
    } catch {
        $e = $_.Exception
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-LogFile -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
        Write-LogFile -LogFile $LogFile -LogString "ERROR: Failed during processing of '$DisplayName' ($UserNameDomainAdmin) - Line $Line" -ForegroundColor Red
        Write-LogFile -LogFile $LogFile -LogString "$e"
        Write-LogFile -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
        Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
        Write-LogFile -LogFile $LogFile -LogString " "
    }
}
#====================================================================
