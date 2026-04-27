#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet("E","H","N")][string]$O365
    , [Parameter(Mandatory)][string]$EmailSuffix
    , [string]$O365EmailSuffix
)

Add-Type -Assembly System.Web

#====================================================================
# Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$AdministrationOU = "Administration"
$SharedAccountsOU = "OU=Shared_Mailbox_Accounts,OU=$AdministrationOU"
$EquipmentAccountsOU = "OU=Equipment_Mailbox_Accounts,OU=$AdministrationOU"
$RoomAccountsOU = "OU=Room_Mailbox_Accounts,OU=$AdministrationOU"
$UsersOU = "Staff"
$O365 = $O365.ToUpper()
if ($O365 -eq "E" -or $O365 -eq "H") {
    if (!$O365EmailSuffix) {
        $O365EmailSuffix = READ-HOST 'Enter "onmicrosoft.com" domain - '
    }
    if ($O365EmailSuffix -notmatch '\.onmicrosoft\.com$') {
        $O365EmailSuffix = "$O365EmailSuffix.onmicrosoft.com"
    }
}
# Group settings
$GroupsOU = "Groups"
$GroupCategory = "Security"
$GroupScope = "Universal"
$SharedGroupsOU = "OU=Shared_Mailbox_Access,OU=$GroupsOU"
$EquipmentGroupsOU = "OU=Equipment_Mailbox_Access,OU=$GroupsOU"
$RoomGroupsOU = "OU=Room_Mailbox_Access,OU=$GroupsOU"
$O365LicenseGroup = "License_Office365"
# ADConnect & Exchange settings
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ExServer = "$Domain-EXCH.$DNSSuffix" #Remote Exchange PS session
$AzureADConnect = "$Domain-RTR.$DNSSuffix"
# Get containing folder for script to locate supporting files
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
# Set variables
$ScriptTitle = "$Domain User Creation Script"
$SMTPServer = $ExServer # SMTP server used for email notifications
$EmailFrom = "noreply@$EmailSuffix" # From address
$PasswordLength = 20 # Number of characters per password
$EnabledMailboxes = @() # Array to Store Completed Mailbox requests for later enumeration
# File locations
$LogPath = "$ScriptPath\LogFiles"
$UserInputFile = "$ScriptPath\users.csv"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
$LogFileName = $Domain + "_new_user_log-$(Get-Date -Format 'yyyyMMdd')"
$LogIndex = 0
while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
    $LogIndex ++
}
$LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
$FailureFile = "$LogPath\$($LogFileName)_$LogIndex-Failures.csv"
#====================================================================

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

#====================================================================
# AD Sync
#====================================================================
function Invoke-ADSync {
    param([pscredential]$Cred)
    try {
        $ADConnectSession = New-PSSession -Computername $AzureADConnect -Credential $Cred
        Invoke-Command -Session $ADConnectSession {Import-Module ADSync}
        Import-PSSession -Session $ADConnectSession -Module ADSync -AllowClobber
        $state = (Get-ADSyncConnectorRunStatus | ? { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
        $ADSyncLoop = 0
        while ($State -and $ADSyncLoop -le 10) {
            Write-Log "AD Sync Connector is currently busy, waiting 30 seconds before trying again"
            Start-Sleep -Seconds 30
            $State = (Get-ADSyncConnectorRunStatus | ? { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
            $ADSyncLoop++
        }
        if ($ADSyncLoop -ge 10) {
            Write-Log "AD Sync Connector has returned a busy state for 5 minutes or more, if this continues, please contact the servicedesk to investigate further"
        } else {
            Write-Log "Attempting to run Azure AD Sync Cycle"
            Start-ADSyncSyncCycle -PolicyType Delta
        }
    } catch {
        $e = $_.Exception
        Write-Log $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log $line -ForegroundColor Red
        $msg = $e.Message
        Write-Log $msg -ForegroundColor Red
        $Action = "Unable to Sync AD"
        Write-Log $Action -ForegroundColor Red
    } finally {
        if ($ADConnectSession) { Remove-PSSession $ADConnectSession }
    }
    $state = (Get-ADSyncConnectorRunStatus | ? { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
    $ADSyncLoop = 0
    while ($State -and $ADSyncLoop -le 10) {
        Write-Log "AD Sync Connector is busy, waiting 30 seconds To allow sync to complete"
        Start-Sleep -Seconds 30
        $State = (Get-ADSyncConnectorRunStatus | ? { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
        $ADSyncLoop++
    }
    if (!($state) -and $ADSyncLoop -le 10) {
        Write-Log "AD Sync complete"
    } else {
        Write-Log "AD Sync has not completed within 5 minutes, please check log for issues relating to syncronization issues."
    }
}
#====================================================================

#====================================================================
# New user email function
#====================================================================
function Send-UserEmail {
    param([string]$UserName,[string]$Requester,[string]$Manager)
    #================================================================
    # Purpose:          To send an email to the requester and/or manager of the new account
    # Assumptions:      Parameters have been set correctly
    # Effects:          Email will be sent to the Requester if field is not blank
    #                   If the manager is different from the requestor, an email
    #                   will also be sent to the manager provided the field is not blank
    # Inputs:           $UserName - SAM account name of user
    #                   $Requester - Person who requested the account, from the CSV
    #                   $Manager - User's manager, as set on the org tab of account properties
    # Calls:            Write-Log function
    # Returns:
    # Notes:
    #================================================================
    # Send email to requester with new user's username & email address
    if ($Requester) {
        $CheckRequester = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$Requester))" -Server $DCHostName
        if ($CheckRequester) {
            Write-Log "Sending email to requester ($Requester) for $UserName..."
            $RequesterEmail = Get-ADUser $Requester -Properties mail | Select-Object -ExpandProperty mail
            $UserEmail = Get-ADUser $UserName -Properties mail | Select-Object -ExpandProperty mail
            $DisplayName = Get-ADUser $UserName -Properties DisplayName | Select-Object -ExpandProperty DisplayName
            $Splat = @{
                To          = $RequesterEmail
                From        = "$ScriptTitle <$EmailFrom>"
                Body        = "New User Created`n`nUserName is $UserName,`nEmail address is $UserEmail.`n`n`nPlease do not reply to this email, it has been sent from an unmonitored address."
                Subject     = "New User Created - $DisplayName"
                SmtpServer  = $SMTPServer
                Priority    = "High"
                UseSsl     = $true
            }
            Send-MailMessage @Splat
        } else {
            Write-Log "WARNING: Cannot send email to requester for $UserName, requester field incorrect..." -ForegroundColor Yellow
        }
    }
    # Send email to manager with new user's username & email address
    if ($Manager) {
        $CheckManager = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$Manager))" -Server $DCHostName
        if ($CheckManager) {
            if ($Manager -ne $Requester) { # check to see if manager is the same as requester, only send email if they're different
                Write-Log "Sending email to manager ($Manager) for $UserName..."
                $ManagerEmail = Get-ADUser $Manager -Properties mail | Select-Object -ExpandProperty mail
                $UserEmail = Get-ADUser $UserName -Properties mail | Select-Object -ExpandProperty mail
                $DisplayName = Get-ADUser $UserName -Properties DisplayName | Select-Object -ExpandProperty DisplayName
                $Splat = @{
                    To          = $ManagerEmail
                    From        = "$ScriptTitle <$EmailFrom>"
                    Body        = "New User Created`n`nUserName is $UserName,`nEmail address is $UserEmail.`n`n`nPlease do not reply to this email, it has been sent from an unmonitored address."
                    Subject     = "New User Created - $DisplayName"
                    SmtpServer  = $SMTPServer
                    Priority    = "High"
                    UseSsl     = $true
                }
                Send-MailMessage @Splat
            }
        } else {
            Write-Log "WARNING: Cannot send email to manager for $UserName, manager field incorrect..." -ForegroundColor Yellow
        }
    } else {
        Write-Log "WARNING: Cannot send email to manager for $UserName, manager field blank..." -ForegroundColor Yellow
    }
}
#====================================================================

#====================================================================
function Send-PasswordEmail {
    param([string]$UserName,[string]$Password,[string]$Requester,[string]$Manager)
    #================================================================
    # Purpose:          To send an email to the requester and/or manager of the new account
    # Assumptions:      Parameters have been set correctly
    # Effects:          Email will be sent to the Requester if field is not blank
    #                   If the manager is different from the requestor, an email
    #                   will also be sent to the manager provided the field is not blank
    # Inputs:           $UserName - SAM account name of user
    #                   $Password - new user password
    #                   $Requester - Person who requested the account, from the CSV
    #                   $Manager - User's manager, as set on the org tab of account properties
    # Calls:            Write-Log function
    # Returns:
    # Notes:
    #================================================================
    # Send email to requester with new user's username & email address
    if ($Requester) {
        $CheckRequester = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$Requester))" -Server $DCHostName
        if ($CheckRequester) {
            Write-Log "Sending email to requester ($Requester) for $UserName..."
            $RequesterEmail = Get-ADUser $Requester -Properties mail | Select-Object -ExpandProperty mail
            $UserEmail = Get-ADUser $UserName -Properties mail | Select-Object -ExpandProperty mail
            $DisplayName = Get-ADUser $UserName -Properties DisplayName | Select-Object -ExpandProperty DisplayName
            $Splat = @{
                To          = $RequesterEmail
                From        = "$ScriptTitle <$EmailFrom>"
                Body        = "New User Created`n`nPassword is $Password.`n`n`nPlease do not reply to this email, it has been sent from an unmonitored address."
                Subject     = "New User Created - $DisplayName"
                SmtpServer  = $SMTPServer
                Priority    = "High"
                UseSsl     = $true
            }
            Send-MailMessage @Splat
        } else {
            Write-Log "WARNING: Cannot send email to requester for $UserName, requester field incorrect..." -ForegroundColor Yellow
        }
    }
    # Send email to manager with new user's username & email address
    if ($Manager) {
        $CheckManager = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$Manager))" -Server $DCHostName
        if ($CheckManager) {
            if ($Manager -ne $Requester) { # check to see if manager is the same as requester, only send email if they're different
                Write-Log "Sending email to manager ($Manager) for $UserName..."
                $ManagerEmail = Get-ADUser $Manager -Properties mail | Select-Object -ExpandProperty mail
                $UserEmail = Get-ADUser $UserName -Properties mail | Select-Object -ExpandProperty mail
                $DisplayName = Get-ADUser $UserName -Properties DisplayName | Select-Object -ExpandProperty DisplayName
                $Splat = @{
                    To          = $ManagerEmail
                    From        = "$ScriptTitle <$EmailFrom>"
                    Body        = "New User Created`n`nPassword is $Password.`n`n`nPlease do not reply to this email, it has been sent from an unmonitored address."
                    Subject     = "New User Created - $DisplayName"
                    SmtpServer  = $SMTPServer
                    Priority    = "High"
                    UseSsl     = $true
                }
                Send-MailMessage @Splat
            }
        } else {
            Write-Log "WARNING: Cannot send email to manager for $UserName, manager field incorrect..." -ForegroundColor Yellow
        }
    } else {
        Write-Log "WARNING: Cannot send email to manager for $UserName, manager field blank..." -ForegroundColor Yellow
    }
}
#====================================================================

#====================================================================
# Start of script
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
$requiredGroups = @('ADM_Task_Standard_Account_Admins', 'ADM_Task_Standard_Group_Admins', 'ADM_Task_SER_Account_Admins', 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

if ($O365 -eq "E" -or $O365 -eq "H") {
    # Get user credentials for server connectivity (Non-MFA)
    try {
        $Cred = Get-Credential -ErrorAction Stop -Message "Admin credentials for remote sessions:"
    } catch {
        $ErrorMsg = $_.Exception.Message
        Write-Log "Failed to validate credentials: $ErrorMsg "
        Read-Host -Prompt "Press Enter to exit"
        Break
    }
    #Connect to remote Exchange PowerShell
    Write-Log "Connecting to remote Exchange PowerShell session... "
    try {
        $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://$ExServer/PowerShell" -Name ExSession -Credential $Cred -ErrorAction Stop -authentication Kerberos
        Write-Log "connected."
        Write-Log "Importing Exchange session... "
        Import-PSSession -Session $ExSession -ErrorAction Stop -AllowClobber > $null
        Write-Log "done."
    } catch {
        $e = $_.Exception
        Write-Log $e
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log $line
        $msg = $e.Message
        Write-Log $msg
        $Action = "Error Importing Exchange Session"
        Write-Log $Action
        Write-Log "failed."
        Write-Log "ERROR: $_" -ForegroundColor Red
    }
    if (!$ExSession) {
        Write-Log "Exchange session not connected Stopping Script"
        Exit
    }
}
#====================================================================

#====================================================================
# Loop through CSV & create users
#====================================================================
# Read input file
Write-Log ("=" * 80)
Write-Log "Reading user data from input file '$UserInputFile'"
Write-Log ("=" * 80)
Write-Log ""
# Read list of users from CSV file ignoring first line
$LIST = @(Import-CSV $UserInputFile)
$RequiredHeaders = @(
    "USERNAME","FIRSTNAME","LASTNAME","DEPT","COMPANY","MANAGER","Requester","S/E/R","AdminID","Managed","Cap","REALNAME","PHONE","HIPRIV","PrivLevel","Description"
)
$Headers = ($LIST | Select-Object -First 1).PSObject.Properties.Name
foreach ($h in $RequiredHeaders) {
    if ($Headers -notcontains $h) {
        throw "CSV missing required column '$h'"
    }
}
$CreatedUsers = @()
# Process each input file record
foreach ($USER in $LIST) {
    $Membership = "$UsersOU"
    $FirstName = $USER.FIRSTNAME
    $FirstName = $FirstName.Trim() -replace '[?@\\+]', [String]::Empty # Strip out illegal characters from First name for Office 365 compliance. Note that \ is escaped to \\
    $LastName = $USER.LASTNAME
    $LastName = $LastName.Trim() -replace '[?@\\+]', [String]::Empty # Strip out illegal characters from LastName for Office 365 compliance. Note that \ is escaped to \\
    $UserName = $USER.USERNAME
    $UserName = $UserName.Trim() -replace '[^A-Za-z0-9.-]', [String]::Empty # Strip out illegal characters from User ID
    if ($UserName.Length -gt 20) {
        $UserName = $UserName.Substring(0,20)
    }
    $Description = $USER.Description
    $Company = $USER.COMPANY
    $Dept = $USER.DEPT
    $HiPriv = $USER.HIPRIV.ToUpper()
    [int]$PrivLevel = $USER.PrivLevel
    $SharedEquipmentRoom = $USER.'S/E/R'.ToUpper()
    $AdminID = $USER.AdminID.ToLower()
    $Managed = $USER.Managed.ToUpper()
    [int]$Capacity = $USER.Cap
    $Phone = $User.PHONE
    $Manager = $USER.MANAGER.ToLower()
    $Requester = $User.Requester.ToLower()
    switch ($SharedEquipmentRoom) {
        "S" {
            $DisplayName = $FirstName
            $LastName = "Shared"
            $OUPath = "$SharedAccountsOU,$EndPath"
            $Enabled = $false
        }
        "E" {
            $DisplayName = $FirstName
            $LastName = "Equipment"
            $OUPath = "$EquipmentAccountsOU,$EndPath"
            $Enabled = $false
        }
        "R" {
            $DisplayName = $FirstName
            $LastName = "Room"
            $OUPath = "$RoomAccountsOU,$EndPath"
            $Enabled = $false
        }
        default {
            $DisplayName = "$LastName, $FirstName"
            $OUPath = "OU=$UsersOU,$EndPath"
            $Enabled = $true
        }
    }
    $RealName = $USER.REALNAME
    if ($RealName) {
        $EmailAddress = "$RealName@$EmailSuffix"
    } else {
        $EmailAddress = "$FirstName.$LastName@$EmailSuffix"
    }
    $HomeDrive = "H:"
    $HomeDir = "\\$DNSSuffix\Profiles\$UserName"
    $UserPrincipalName = "$UserName@$EmailSuffix"
    $ExistingUser = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$UserName))" -Server $DCHostName
    if ($ExistingUser) {
        Write-Log ("-" * 80) -ForegroundColor Red
        Write-Log "ERROR: User '$($UserName)' already exists in the $Domain directory. The user object`n is '$($ExistingUser.DistinguishedName)'" -ForegroundColor Red
        Write-Log "Error: Processing input file record for '$LastName, $FirstName' ($UserName) aborted" -ForegroundColor Red
        Write-Log ("-" * 80) -ForegroundColor Red
        Write-Log ""
        continue # Skip this user
    } else {
        try {
            Write-Log ("=" * 80)
            Write-Log "Processing input file record for '$DisplayName' ($UserName)..."
            Write-Log ("=" * 80)
            # Generate random password
            $UserPassword = [Web.Security.Membership]::GeneratePassword($PasswordLength,4)
            # Test password against password policy
            Test-Password -Password $UserPassword
            $Params = @{
                Name                    = $UserName
                AccountPassword         = ConvertTo-SecureString -AsPlainText $UserPassword -Force
                ChangePasswordAtLogon   = $true
                Company                 = $Company
                Department              = $Dept
                Description             = $Description
                DisplayName             = $DisplayName
                EmailAddress            = $EmailAddress
                Enabled                 = $Enabled
                GivenName               = $FirstName
                HomeDirectory           = $HomeDir
                HomeDrive               = $HomeDrive
                OfficePhone             = $Phone
                ProfilePath             = $HomeDir
                Path                    = $OUPath
                SamAccountName          = $UserName
                SurName                 = $LastName
                UserPrincipalName       = $UserPrincipalName
            }
            Write-Log "Creating $UserName"
            New-ADUser -Type "user" -Server $DCHostName @Params
            Write-Log "Created $UserName"
            Set-ADAccountControl -AccountNotDelegated $false -AllowReversiblePasswordEncryption $false -CannotChangePassword $false -DoesNotRequirePreAuth $false -Identity "CN=$UserName,$OUPath" -PasswordNeverExpires $false -UseDESKeyOnly $false -Server $DCHostName
            Add-GroupMember -Group $Membership -Member $UserName
            if ($Enabled -and $HomeDir) {
                $null = New-Item -Path $HomeDir -ItemType Directory -Force -ErrorAction SilentlyContinue
                try {
                    $Acl = Get-Acl $HomeDir
                    $Ar = New-Object System.Security.AccessControl.FileSystemAccessRule("$Domain\$UserName","Modify","ContainerInherit,ObjectInherit","None","Allow")
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
            }
            if ($Dept) {
                Add-GroupMember -Group $Dept -Member $UserName
            }
            switch ($SharedEquipmentRoom) {
                "S" {
                    # create management group for shared account
                    $GroupName = "sh_$UserName"
                    Write-Log "Creating group $GroupName for shared account management"
                    New-DomainGroup -GroupName $GroupName -GroupScope $GroupScope -O365 $O365 -HiddenFromAddressListsEnabled $true -Path "$SharedGroupsOU,$EndPath" -GroupDescription "Group to grant access access to $UserName"
                    Write-Log "Group $GroupName created in location $SharedGroupsOU,$EndPath"
                    Set-ADObject -Identity "CN=$UserName,$OUPath" -protectedFromAccidentalDeletion $True -Server $DCHostName
                    if ($AdminID) {
                        $CheckAdminID = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$AdminID))" -Server $DCHostName
                        if ($CheckAdminID) {
                            Add-GroupMember -Group $GroupName -Member $AdminID
                        }
                    }
                }
                "E" {
                    # create management group for equipment account
                    $GroupName = "eq_$UserName"
                    Write-Log "Creating group $GroupName for equipment account management"
                    New-DomainGroup -GroupName $GroupName -GroupScope $GroupScope -O365 $O365 -HiddenFromAddressListsEnabled $true -Path "$EquipmentGroupsOU,$EndPath" -GroupDescription "Group to grant access access to $UserName"
                    Write-Log "Group $GroupName created in location $EquipmentGroupsOU,$EndPath"
                    Set-ADObject -Identity "CN=$UserName,$OUPath" -protectedFromAccidentalDeletion $True -Server $DCHostName
                    if ($Capacity) {
                        Write-Log "Setting Title for Equipment Account to 'Cap: $Capacity'..."
                        Set-ADUser -Identity $UserName -Title "Cap: $Capacity" -Server $DCHostName
                    }
                    if ($AdminID) {
                        if ($Managed -eq "M") {
                            $Assistant = $AdminID + " (M)"
                        } else {
                            $Assistant = $AdminID
                        }
                        Write-Log "Setting Assistant for Equipment Account to '$Assistant'..."
                        Set-ADUser -Identity $UserName -Replace @{msExchAssistantName=$Assistant} -Server $DCHostName
                        $CheckAdminID = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$AdminID))" -Server $DCHostName
                        if ($CheckAdminID) {
                            Add-GroupMember -Group $GroupName -Member $AdminID
                        }
                    }
                }
                "R" {
                    # create management group for room account
                    $GroupName = "ro_$UserName"
                    Write-Log "Creating group $GroupName for room account management"
                    New-DomainGroup -GroupName $GroupName -GroupScope $GroupScope -O365 $O365 -HiddenFromAddressListsEnabled $true -Path "$RoomGroupsOU,$EndPath" -GroupDescription "Group to grant access access to $UserName"
                    Write-Log "Group $GroupName created in location $RoomGroupsOU,$EndPath"
                    Set-ADObject -Identity "CN=$UserName,$OUPath" -protectedFromAccidentalDeletion $True -Server $DCHostName
                    if ($Capacity) {
                        Write-Log "Setting Title for Room Account to 'Cap: $Capacity'..."
                        Set-ADUser -Identity $UserName -Title "Cap: $Capacity" -Server $DCHostName
                    }
                    if ($AdminID) {
                        if ($Managed -eq "M") {
                            $Assistant = $AdminID + " (M)"
                        } else {
                            $Assistant = $AdminID
                        }
                        Write-Log "Setting Assistant for Room Account to '$Assistant'..."
                        Set-ADUser -Identity $UserName -Replace @{msExchAssistantName=$Assistant} -Server $DCHostName
                        $CheckAdminID = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$AdminID))" -Server $DCHostName
                        if ($CheckAdminID) {
                            Add-GroupMember -Group $GroupName -Member $AdminID
                        }
                    }
                }
                default {
                    if ($O365LicenseGroup) {
                        Add-GroupMember -Group $O365LicenseGroup -Member $UserName
                    }
                    # Set manager
                    if ($Manager) {
                        $CheckManager = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$Manager))" -Server $DCHostName
                        if ($CheckManager) {
                            Write-Log "Setting manager for $UserName to $Manager..."
                            Set-ADUser -Identity $UserName -Manager $Manager -Server $DCHostName
                        } else {
                            Write-Log "WARNING: Cannot set manager for $UserName, manager field incorrect..." -ForegroundColor Yellow
                        }
                    } else {
                        Write-Log "WARNING: Cannot set manager for $UserName, manager field blank..." -ForegroundColor Yellow
                    }
                    if ($Dept -eq "IT" -and $HiPriv -eq "Y") {
                        Add-GroupMember -Group "sh_ITHELP" -Member $UserName
                        Write-Log "Creating HiPriv account for $UserName"
                        Write-Log ""
                        & $PSScriptRoot\CreateITAdminUser.ps1 -FirstName $FirstName -LastName $LastName -UserName $UserName -EmailSuffix $EmailSuffix -PrivLevel $PrivLevel -Description $Description -Dept $Dept -Company $Company -LogFile $LogFile -DCHostName $DCHostName -Manager $UserName -PasswordLength $PasswordLength
                    }
                }
            }
            if ($O365 -eq "E" -or $O365 -eq "H") {
                if ($O365 -eq "E") {
                    Write-Log "Exchange mailbox for $UserName will be created in Exchange OnPrem"
                    Write-Log "Calling New-UserOnPremMailbox function with the following parameters:"
                    Write-Log "UserName: $UserName, realname: $RealName, SharedEquipmentRoom: $SharedEquipmentRoom, Capacity: $Capacity"
                    $EnabledMailboxes += New-UserOnPremMailbox -UserName $UserName -realname $RealName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
                }
                if ($O365 -eq "H") {
                    Write-Log "Exchange mailbox for $UserName will be created in Exchange Online"
                    Write-Log "Calling New-UserMailbox function with the following parameters:"
                    Write-Log "UserName: $UserName, realname: $RealName, SharedEquipmentRoom: $SharedEquipmentRoom, Capacity: $Capacity"
                    $EnabledMailboxes += New-UserMailbox -UserName $UserName -realname $RealName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
                }
                # Send email to requester and manager with new user's username & email address
                if (-not $SharedEquipmentRoom) {
                    Send-UserEmail -UserName $UserName -requester $Requester -manager $Manager
                    Send-PasswordEmail -UserName $UserName -Password $UserPassword -requester $Requester -manager $Manager
                }
            }
            $UserPassword = $null
            $CreatedUsers += $USER
            Write-Log ("=" * 80)
            Write-Log "Processing input file record for '$DisplayName' ($UserName) complete"
            Write-Log ("=" * 80)
            Write-Log ""
        } catch {
            $e = $_.Exception
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log ("-" * 80) -ForegroundColor Red
            Write-Log "ERROR: Failed during processing of '$DisplayName' ($UserName) - Line $Line" -ForegroundColor Red
            Write-Log "$e"
            Write-Log ("-" * 80) -ForegroundColor Red
            Write-Log ("=" * 80)
            Write-Log ""
        }
    }
}
if ($O365 -eq "E") {
    foreach ($mailbox in $EnabledMailboxes) {
        $i = 0
        $MBX = $null
        Do {
            $MBX = Get-Mailbox $mailbox.alias -Server $DCHostName -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 30
            $i++
        } While (!($MBX) -and $i -lt 5)
        if ($MBX) {
            if ($Mailbox.SharedEquipmentRoom) {
                $logmsg = "Updating Mailbox: " + $Mailbox.Alias +" "+ $Mailbox.SharedEquipmentRoom +" "+ $Mailbox.Capacity +" "
                Write-Log $logMsg
                Update-UserOnPremMailbox $mailbox.Alias $Mailbox.SharedEquipmentRoom $Mailbox.Capacity
            } elseif (!$Mailbox.SharedEquipmentRoom -and !$Mailbox.Capacity) {
                $logmsg = "Updating Mailbox: " + $Mailbox.Alias
                Write-Log $logMsg
                Update-UserOnPremMailbox $mailbox.Alias $Mailbox.SharedEquipmentRoom $Mailbox.Capacity
            }
        } else {
            $logmsg = "Mailbox: " + $Mailbox.Alias +" not found in AD"
            Write-Log $logMsg
        }
    }
    if (Get-PSSession -Name ExSession -ErrorAction SilentlyContinue) {
        Remove-PSSession $ExSession
        Write-Log "Closed Exchange session."
    }
}
if ($O365 -eq "H") {
    if (Get-PSSession -Name ExSession -ErrorAction SilentlyContinue) {
        Remove-PSSession $ExSession
        Write-Log "Closed Exchange session."
    }
    Get-PSSession | Remove-PSSession
    #Force ADSync
    Invoke-ADSync -Cred $Cred
    $Connected = $false
    $Failures = @()
    if ($Connected -eq $false) {
        if (!(Get-Module -ListAvailable -Name Microsoft.Graph)) {
            Write-Log "Microsoft.Graph module not installed"
            throw "Microsoft.Graph module not installed"
        }
        if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Log "ExchangeOnlineManagement module not installed"
            throw "ExchangeOnlineManagement module not installed"
        }
        try {
            Write-Log "Connecting to Exchange Online"
            Import-Module -Name ExchangeOnlineManagement
            Connect-ExchangeOnline
            Write-Log "Connected to Exchange Online"
            Write-Log "Connecting to Microsoft Graph"
            Import-Module -Name Microsoft.Graph.Authentication
            Import-Module -Name Microsoft.Graph.Users
            Import-Module -Name Microsoft.Graph.Identity.Governance
            Connect-MgGraph -NoWelcome -Scopes "RoleManagement.ReadWrite.Directory", "User.ReadWrite.All"
            Write-Log "Connected to Microsoft Graph"
            $Connected = $true
        } catch {
            $e = $_.Exception
            Write-Log $e
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log $line
            $msg = $e.Message
            Write-Log $msg
            $Action = "Failed to connect to Exchange Online on import session"
            Write-Log $Action
        }
    }
    if ($Connected -eq $true) {
        Write-Log "Updating Mailboxes"
        $LastTry = $True
        Foreach ($mailbox in $EnabledMailboxes) {
            $i = 0
            $MBX = $null
            Do {
                $MBX = Get-Mailbox $mailbox.alias -Server $DCHostName -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 30
                $i++
                if ($i -eq 4 -and $LastTry -eq $True) {
                    Invoke-ADSync -Cred $Cred
                    $LastTry = $False
                    $i = 0
                }
            } While (!($MBX) -and $i -lt 5)
            if ($MBX) {
                if ($Mailbox.SharedEquipmentRoom) {
                    $logmsg = "Updating Mailbox: " + $Mailbox.Alias +" "+ $Mailbox.SharedEquipmentRoom +" "+ $Mailbox.Capacity +" "+ $Mailbox.AdminID
                    Write-Log $logMsg
                    Update-UserMailbox $mailbox.Alias $Mailbox.SharedEquipmentRoom $Mailbox.Capacity
                } elseif (!$Mailbox.SharedEquipmentRoom -and !$Mailbox.Capacity) {
                    $logmsg = "Updating Mailbox: " + $Mailbox.Alias
                    Write-Log $logMsg
                    Update-UserMailbox $mailbox.Alias $Mailbox.SharedEquipmentRoom $Mailbox.Capacity
                }
            } else {
                $logmsg = "Mailbox: " + $Mailbox.Alias +" not found in AzureAD"
                $Failures += $Mailbox
                Write-Log $logMsg
            }
        }
        foreach ($USER in $CreatedUsers) {
            try {
                $FirstName = $USER.FIRSTNAME
                $FirstName = $FirstName.Trim() -replace '[?@\\+]', [String]::Empty # Strip out illegal characters from First name for Office 365 compliance. Note that \ is escaped to \\
                $LastName = $USER.LASTNAME
                $LastName = $LastName.Trim() -replace '[?@\\+]', [String]::Empty # Strip out illegal characters from LastName for Office 365 compliance. Note that \ is escaped to \\
                $UserName = $USER.USERNAME
                $UserName = $UserName.Trim() -replace '[^A-Za-z0-9.-]', [String]::Empty # Strip out illegal characters from User ID
                if ($UserName.Length -gt 20) {
                    $UserName = $UserName.Substring(0,20)
                }
                $Company = $USER.COMPANY
                $Dept = $USER.DEPT
                $HiPriv = $USER.HIPRIV.ToUpper()
                [int]$PrivLevel = $USER.PrivLevel
                $SharedEquipmentRoom = $USER.'S/E/R'.ToUpper()
                [int]$Capacity = $USER.Cap
                $UserPrincipalName = "$UserName@$EmailSuffix"
                Write-Log ""
                Write-Log "Assigning region for $UserName"
                Update-MgUser -UserId $UserPrincipalName -UsageLocation GB
                switch ($SharedEquipmentRoom) {
                    default {
                        if ($Dept -eq "IT" -and $HiPriv -eq "Y") {
                            Write-Log "Creating Cloud Admin account for $UserName"
                            Write-Log ""
                            & $PSScriptRoot\CreateITCloudAdminUser.ps1 -FirstName $FirstName -LastName $LastName -UserName $UserName -EmailSuffix $EmailSuffix -PrivLevel $PrivLevel -Dept $Dept -Company $Company -LogFile $LogFile -Manager $UserName -PasswordLength $PasswordLength
                        }
                    }
                }
            } catch {
                $e = $_.Exception
                $line = $_.InvocationInfo.ScriptLineNumber
                Write-Log ("-" * 80) -ForegroundColor Red
                Write-Log "ERROR processing '$UserName' - Line $line : $($e.Message)" -ForegroundColor Red
                Write-Log ("-" * 80) -ForegroundColor Red
                continue
            }
        }
        Disconnect-ExchangeOnline -Confirm:$false
        Disconnect-MgGraph
    }
    if (Get-PSSession) {
        Write-Log "Cleaning up PSSessions"
        Get-PSSession | Remove-PSSession
    }
    if ($Failures) {
        if ($Failures.Count -gt 0) {
            Foreach ($Failure in $Failures) {
                $LIST | Where-Object {
                    ($_.USERNAME.Trim() -replace '[^A-Za-z0-9.-]','') -ieq $Failure.alias
                } | Export-Csv $FailureFile -NoTypeInformation -Append
            }
        }
    }
}
Write-Log ("=" * 80)
Write-Log "Processing complete"
Write-Log ("=" * 80)
#====================================================================
