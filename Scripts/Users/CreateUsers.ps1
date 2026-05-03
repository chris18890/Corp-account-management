#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet("E","H","N")][string]$O365
    , [Parameter(Mandatory)][string]$EmailSuffix
    , [string]$O365EmailSuffix
)

Set-StrictMode -Version Latest
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
    param([pscredential]$Cred,[String]$AzureADConnect,[String]$O365EmailSuffix)
    try {
        $ADConnectSession = New-PSSession -Computername $AzureADConnect -Credential $Cred
        Invoke-Command -Session $ADConnectSession {Import-Module ADSync}
        Import-PSSession -Session $ADConnectSession -Module ADSync -AllowClobber
        $state = (Get-ADSyncConnectorRunStatus | Where-Object { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
        $ADSyncLoop = 0
        while ($State -and $ADSyncLoop -le 10) {
            Write-Log -LogFile $LogFile -LogString "AD Sync Connector is currently busy, waiting 30 seconds before trying again"
            Start-Sleep -Seconds 30
            $State = (Get-ADSyncConnectorRunStatus | Where-Object { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
            $ADSyncLoop++
        }
        if ($ADSyncLoop -ge 10) {
            Write-Log -LogFile $LogFile -LogString "AD Sync Connector has returned a busy state for 5 minutes or more, if this continues, please contact the servicedesk to investigate further"
        } else {
            Write-Log -LogFile $LogFile -LogString "Attempting to run Azure AD Sync Cycle"
            Start-ADSyncSyncCycle -PolicyType Delta
        }
        $state = (Get-ADSyncConnectorRunStatus | Where-Object { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
        $ADSyncLoop = 0
        while ($State -and $ADSyncLoop -le 10) {
            Write-Log -LogFile $LogFile -LogString "AD Sync Connector is busy, waiting 30 seconds To allow sync to complete"
            Start-Sleep -Seconds 30
            $State = (Get-ADSyncConnectorRunStatus | Where-Object { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
            $ADSyncLoop++
        }
        if (!($state) -and $ADSyncLoop -le 10) {
            Write-Log -LogFile $LogFile -LogString "AD Sync complete"
        } else {
            Write-Log -LogFile $LogFile -LogString "AD Sync has not completed within 5 minutes, please check log for issues relating to syncronization issues."
        }
    } catch {
        $e = $_.Exception
        Write-Log -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-Log -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Unable to Sync AD"
        Write-Log -LogFile $LogFile -LogString $Action -ForegroundColor Red
    } finally {
        if ($ADConnectSession) { Remove-PSSession $ADConnectSession }
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
            Write-Log -LogFile $LogFile -LogString "Sending email to requester ($Requester) for $UserName..."
            $RequesterEmail = Get-ADUser $Requester -Properties mail -Server $DCHostName | Select-Object -ExpandProperty mail
            $UserEmail = Get-ADUser $UserName -Properties mail -Server $DCHostName | Select-Object -ExpandProperty mail
            $DisplayName = Get-ADUser $UserName -Properties DisplayName -Server $DCHostName | Select-Object -ExpandProperty DisplayName
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
            Write-Log -LogFile $LogFile -LogString "WARNING: Cannot send email to requester for $UserName, requester field incorrect..." -ForegroundColor Yellow
        }
    }
    # Send email to manager with new user's username & email address
    if ($Manager) {
        $CheckManager = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$Manager))" -Server $DCHostName
        if ($CheckManager) {
            if ($Manager -ne $Requester) { # check to see if manager is the same as requester, only send email if they're different
                Write-Log -LogFile $LogFile -LogString "Sending email to manager ($Manager) for $UserName..."
                $ManagerEmail = Get-ADUser $Manager -Properties mail -Server $DCHostName | Select-Object -ExpandProperty mail
                $UserEmail = Get-ADUser $UserName -Properties mail -Server $DCHostName | Select-Object -ExpandProperty mail
                $DisplayName = Get-ADUser $UserName -Properties DisplayName -Server $DCHostName | Select-Object -ExpandProperty DisplayName
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
            Write-Log -LogFile $LogFile -LogString "WARNING: Cannot send email to manager for $UserName, manager field incorrect..." -ForegroundColor Yellow
        }
    } else {
        Write-Log -LogFile $LogFile -LogString "WARNING: Cannot send email to manager for $UserName, manager field blank..." -ForegroundColor Yellow
    }
}
#====================================================================

#====================================================================
# Start of script
#====================================================================
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-Log -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-Log -LogFile $LogFile -LogString "$ScriptTitle"
Write-Log -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString " "
$requiredGroups = @('ADM_Task_Standard_Account_Admins', 'ADM_Task_Standard_Group_Admins', 'ADM_Task_SER_Account_Admins', 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

if ($O365 -eq "E" -or $O365 -eq "H") {
    # Get user credentials for server connectivity (Non-MFA)
    try {
        $Cred = Get-Credential -ErrorAction Stop -Message "Admin credentials for remote sessions:"
    } catch {
        $ErrorMsg = $_.Exception.Message
        Write-Log -LogFile $LogFile -LogString "Failed to validate credentials: $ErrorMsg "
        Read-Host -Prompt "Press Enter to exit"
        Exit
    }
    #Connect to remote Exchange PowerShell
    Write-Log -LogFile $LogFile -LogString "Connecting to remote Exchange PowerShell session... "
    try {
        $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://$ExServer/PowerShell" -Name ExSession -Credential $Cred -ErrorAction Stop -authentication Kerberos
        Write-Log -LogFile $LogFile -LogString "connected."
        Write-Log -LogFile $LogFile -LogString "Importing Exchange session... "
        Import-PSSession -Session $ExSession -ErrorAction Stop -AllowClobber > $null
        Write-Log -LogFile $LogFile -LogString "done."
    } catch {
        $e = $_.Exception
        Write-Log -LogFile $LogFile -LogString $e
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log -LogFile $LogFile -LogString $line
        $msg = $e.Message
        Write-Log -LogFile $LogFile -LogString $msg
        $Action = "Error Importing Exchange Session"
        Write-Log -LogFile $LogFile -LogString $Action
        Write-Log -LogFile $LogFile -LogString "failed."
        Write-Log -LogFile $LogFile -LogString "ERROR: $_" -ForegroundColor Red
    }
    if (!$ExSession) {
        Write-Log -LogFile $LogFile -LogString "Exchange session not connected Stopping Script"
        Exit
    }
}
#====================================================================

#====================================================================
# Loop through CSV & create users
#====================================================================
# Read input file
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString "Reading user data from input file '$UserInputFile'"
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString " "
# Read list of users from CSV file ignoring first line
$UserList = @(Import-CSV $UserInputFile)
if ($UserList -isnot [Array]) {$UserList = @($UserList)}
$RequiredHeaders = @(
    "USERNAME","FIRSTNAME","LASTNAME","DEPT","COMPANY","MANAGER","Requester","S/E/R","AdminID","Managed","Cap","REALNAME","PHONE","HIPRIV","PrivLevel","Description"
)
$Headers = ($UserList | Select-Object -First 1).PSObject.Properties.Name
foreach ($h in $RequiredHeaders) {
    if ($Headers -notcontains $h) {
        throw "CSV missing required column '$h'"
    }
}
$CreatedUsers = @()
# Process each input file record
foreach ($USER in $UserList) {
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
        Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
        Write-Log -LogFile $LogFile -LogString "ERROR: User '$($UserName)' already exists in the $Domain directory. The user object`n is '$($ExistingUser.DistinguishedName)'" -ForegroundColor Red
        Write-Log -LogFile $LogFile -LogString "Error: Processing input file record for '$LastName, $FirstName' ($UserName) aborted" -ForegroundColor Red
        Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
        Write-Log -LogFile $LogFile -LogString " "
        continue # Skip this user
    } else {
        try {
            Write-Log -LogFile $LogFile -LogString ("=" * 80)
            Write-Log -LogFile $LogFile -LogString "Processing input file record for '$DisplayName' ($UserName)..."
            Write-Log -LogFile $LogFile -LogString ("=" * 80)
            # Generate random password
            $UserPassword = [Web.Security.Membership]::GeneratePassword($PasswordLength,4)
            # Test password against password policy
            Test-Password -LogFile $LogFile -Password $UserPassword -PasswordLength $PasswordLength
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
            Write-Log -LogFile $LogFile -LogString "Creating $UserName"
            New-ADUser -Type "user" -Server $DCHostName @Params
            $UserPassword = $null
            Write-Log -LogFile $LogFile -LogString "Created $UserName"
            Set-ADAccountControl -AccountNotDelegated $false -AllowReversiblePasswordEncryption $false -CannotChangePassword $false -DoesNotRequirePreAuth $false -Identity "CN=$UserName,$OUPath" -PasswordNeverExpires $false -UseDESKeyOnly $false -Server $DCHostName
            Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $Membership -Member $UserName
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
                    Write-Log -LogFile $LogFile -LogString "Created home directory $HomeDir"
                } catch {
                    Write-Log -LogFile $LogFile -LogString "WARNING: Could not set ACL on $HomeDir - $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            if ($Dept) {
                Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $Dept -Member $UserName
            }
            switch ($SharedEquipmentRoom) {
                "S" {
                    # create management group for shared account
                    $GroupName = "sh_$UserName"
                    Write-Log -LogFile $LogFile -LogString "Creating group $GroupName for shared account management"
                    New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $GroupName -GroupCategory $GroupCategory -GroupScope $GroupScope -O365 $O365 -HiddenFromAddressListsEnabled $true -Path "$SharedGroupsOU,$EndPath" -GroupDescription "Group to grant access access to $UserName"
                    Write-Log -LogFile $LogFile -LogString "Group $GroupName created in location $SharedGroupsOU,$EndPath"
                    Set-ADObject -Identity "CN=$UserName,$OUPath" -protectedFromAccidentalDeletion $True -Server $DCHostName
                    if ($AdminID) {
                        $CheckAdminID = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$AdminID))" -Server $DCHostName
                        if ($CheckAdminID) {
                            Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $GroupName -Member $AdminID
                        }
                    }
                }
                "E" {
                    # create management group for equipment account
                    $GroupName = "eq_$UserName"
                    Write-Log -LogFile $LogFile -LogString "Creating group $GroupName for equipment account management"
                    New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $GroupName -GroupCategory $GroupCategory -GroupScope $GroupScope -O365 $O365 -HiddenFromAddressListsEnabled $true -Path "$EquipmentGroupsOU,$EndPath" -GroupDescription "Group to grant access access to $UserName"
                    Write-Log -LogFile $LogFile -LogString "Group $GroupName created in location $EquipmentGroupsOU,$EndPath"
                    Set-ADObject -Identity "CN=$UserName,$OUPath" -protectedFromAccidentalDeletion $True -Server $DCHostName
                    if ($Capacity) {
                        Write-Log -LogFile $LogFile -LogString "Setting Title for Equipment Account to 'Cap: $Capacity'..."
                        Set-ADUser -Identity $UserName -Title "Cap: $Capacity" -Server $DCHostName
                    }
                    if ($AdminID) {
                        if ($Managed -eq "M") {
                            $Assistant = $AdminID + " (M)"
                        } else {
                            $Assistant = $AdminID
                        }
                        Write-Log -LogFile $LogFile -LogString "Setting Assistant for Equipment Account to '$Assistant'..."
                        Set-ADUser -Identity $UserName -Replace @{msExchAssistantName=$Assistant} -Server $DCHostName
                        $CheckAdminID = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$AdminID))" -Server $DCHostName
                        if ($CheckAdminID) {
                            Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $GroupName -Member $AdminID
                        }
                    }
                }
                "R" {
                    # create management group for room account
                    $GroupName = "ro_$UserName"
                    Write-Log -LogFile $LogFile -LogString "Creating group $GroupName for room account management"
                    New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $GroupName -GroupCategory $GroupCategory -GroupScope $GroupScope -O365 $O365 -HiddenFromAddressListsEnabled $true -Path "$RoomGroupsOU,$EndPath" -GroupDescription "Group to grant access access to $UserName"
                    Write-Log -LogFile $LogFile -LogString "Group $GroupName created in location $RoomGroupsOU,$EndPath"
                    Set-ADObject -Identity "CN=$UserName,$OUPath" -protectedFromAccidentalDeletion $True -Server $DCHostName
                    if ($Capacity) {
                        Write-Log -LogFile $LogFile -LogString "Setting Title for Room Account to 'Cap: $Capacity'..."
                        Set-ADUser -Identity $UserName -Title "Cap: $Capacity" -Server $DCHostName
                    }
                    if ($AdminID) {
                        if ($Managed -eq "M") {
                            $Assistant = $AdminID + " (M)"
                        } else {
                            $Assistant = $AdminID
                        }
                        Write-Log -LogFile $LogFile -LogString "Setting Assistant for Room Account to '$Assistant'..."
                        Set-ADUser -Identity $UserName -Replace @{msExchAssistantName=$Assistant} -Server $DCHostName
                        $CheckAdminID = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$AdminID))" -Server $DCHostName
                        if ($CheckAdminID) {
                            Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $GroupName -Member $AdminID
                        }
                    }
                }
                default {
                    if ($O365LicenseGroup) {
                        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $O365LicenseGroup -Member $UserName
                    }
                    # Set manager
                    if ($Manager) {
                        $CheckManager = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$Manager))" -Server $DCHostName
                        if ($CheckManager) {
                            Write-Log -LogFile $LogFile -LogString "Setting manager for $UserName to $Manager..."
                            Set-ADUser -Identity $UserName -Manager $Manager -Server $DCHostName
                        } else {
                            Write-Log -LogFile $LogFile -LogString "WARNING: Cannot set manager for $UserName, manager field incorrect..." -ForegroundColor Yellow
                        }
                    } else {
                        Write-Log -LogFile $LogFile -LogString "WARNING: Cannot set manager for $UserName, manager field blank..." -ForegroundColor Yellow
                    }
                    if ($Dept -eq "IT" -and $HiPriv -eq "Y") {
                        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "sh_ITHELP" -Member $UserName
                        Write-Log -LogFile $LogFile -LogString "Creating HiPriv account for $UserName"
                        Write-Log -LogFile $LogFile -LogString " "
                        & $PSScriptRoot\CreateITAdminUser.ps1 -FirstName $FirstName -LastName $LastName -UserName $UserName -EmailSuffix $EmailSuffix -PrivLevel $PrivLevel -Description $Description -Dept $Dept -Company $Company -LogFile $LogFile -DCHostName $DCHostName -Manager $UserName -PasswordLength $PasswordLength
                    }
                }
            }
            if ($O365 -eq "E" -or $O365 -eq "H") {
                if ($O365 -eq "E") {
                    Write-Log -LogFile $LogFile -LogString "Exchange mailbox for $UserName will be created in Exchange OnPrem"
                    Write-Log -LogFile $LogFile -LogString "Calling New-UserOnPremMailbox function with the following parameters:"
                    Write-Log -LogFile $LogFile -LogString "UserName: $UserName, EmailSuffix: $EmailSuffix, realname: $RealName, SharedEquipmentRoom: $SharedEquipmentRoom, Capacity: $Capacity"
                    $EnabledMailboxes += New-UserOnPremMailbox -LogFile $LogFile -DCHostName $DCHostName -UserName $UserName -EmailSuffix $EmailSuffix -realname $RealName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
                }
                if ($O365 -eq "H") {
                    Write-Log -LogFile $LogFile -LogString "Exchange mailbox for $UserName will be created in Exchange Online"
                    Write-Log -LogFile $LogFile -LogString "Calling New-UserMailbox function with the following parameters:"
                    Write-Log -LogFile $LogFile -LogString "UserName: $UserName, EmailSuffix: $EmailSuffix, realname: $RealName, SharedEquipmentRoom: $SharedEquipmentRoom, Capacity: $Capacity"
                    $EnabledMailboxes += New-UserMailbox -LogFile $LogFile -DCHostName $DCHostName -UserName $UserName -EmailSuffix $EmailSuffix -O365EmailSuffix $O365EmailSuffix -realname $RealName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
                }
                # Send email to requester and manager with new user's username & email address
                if (-not $SharedEquipmentRoom) {
                    Send-UserEmail -UserName $UserName -requester $Requester -manager $Manager
                }
            }
            $CreatedUsers += $USER
            Write-Log -LogFile $LogFile -LogString ("=" * 80)
            Write-Log -LogFile $LogFile -LogString "Processing input file record for '$DisplayName' ($UserName) complete"
            Write-Log -LogFile $LogFile -LogString ("=" * 80)
            Write-Log -LogFile $LogFile -LogString " "
        } catch {
            $e = $_.Exception
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
            Write-Log -LogFile $LogFile -LogString "ERROR: Failed during processing of '$DisplayName' ($UserName) - Line $Line" -ForegroundColor Red
            Write-Log -LogFile $LogFile -LogString "$e"
            Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
            Write-Log -LogFile $LogFile -LogString ("=" * 80)
            Write-Log -LogFile $LogFile -LogString " "
        }
    }
}
if ($O365 -eq "E") {
    foreach ($mailbox in $EnabledMailboxes) {
        $i = 0
        $MBX = $null
        Do {
            $MBX = Get-Mailbox -Identity $mailbox.alias -DomainController $DCHostName -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 30
            $i++
        } While (!($MBX) -and $i -lt 5)
        if ($MBX) {
            if ($Mailbox.SharedEquipmentRoom) {
                $logmsg = "Updating Mailbox: " + $Mailbox.Alias +" "+ $Mailbox.SharedEquipmentRoom +" "+ $Mailbox.Capacity +" "
                Write-Log -LogFile $LogFile -LogString $LogMsg
            } elseif (!$Mailbox.SharedEquipmentRoom -and !$Mailbox.Capacity) {
                $logmsg = "Updating Mailbox: " + $Mailbox.Alias
                Write-Log -LogFile $LogFile -LogString $LogMsg
            }
            Update-UserOnPremMailbox -LogFile $LogFile -DCHostName $DCHostName -UserName $mailbox.Alias -SharedEquipmentRoom $Mailbox.SharedEquipmentRoom -Capacity $Mailbox.Capacity
        } else {
            $logmsg = "Mailbox: " + $Mailbox.Alias +" not found in AD"
            Write-Log -LogFile $LogFile -LogString $LogMsg
        }
    }
    if (Get-PSSession -Name ExSession -ErrorAction SilentlyContinue) {
        Remove-PSSession $ExSession
        $Cred.Dispose()
        Write-Log -LogFile $LogFile -LogString "Closed Exchange session."
    }
}
if ($O365 -eq "H") {
    if (Get-PSSession -Name ExSession -ErrorAction SilentlyContinue) {
        Remove-PSSession $ExSession
        Write-Log -LogFile $LogFile -LogString "Closed Exchange session."
    }
    Get-PSSession | Remove-PSSession
    #Force ADSync
    Invoke-ADSync -Cred $Cred -AzureADConnect $AzureADConnect -O365EmailSuffix $O365EmailSuffix
    $Connected = $false
    $Failures = @()
    if ($Connected -eq $false) {
        if (!(Get-Module -ListAvailable -Name Microsoft.Graph)) {
            Write-Log -LogFile $LogFile -LogString "Microsoft.Graph module not installed"
            throw "Microsoft.Graph module not installed"
        }
        if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Log -LogFile $LogFile -LogString "ExchangeOnlineManagement module not installed"
            throw "ExchangeOnlineManagement module not installed"
        }
        try {
            Write-Log -LogFile $LogFile -LogString "Connecting to Exchange Online"
            Import-Module -Name ExchangeOnlineManagement
            Connect-ExchangeOnline
            Write-Log -LogFile $LogFile -LogString "Connected to Exchange Online"
            Write-Log -LogFile $LogFile -LogString "Connecting to Microsoft Graph"
            Import-Module -Name Microsoft.Graph.Authentication
            Import-Module -Name Microsoft.Graph.Users
            Import-Module -Name Microsoft.Graph.Identity.Governance
            Connect-MgGraph -NoWelcome -Scopes "RoleManagement.ReadWrite.Directory", "User.ReadWrite.All"
            Write-Log -LogFile $LogFile -LogString "Connected to Microsoft Graph"
            $Connected = $true
        } catch {
            $e = $_.Exception
            Write-Log -LogFile $LogFile -LogString $e
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log -LogFile $LogFile -LogString $line
            $msg = $e.Message
            Write-Log -LogFile $LogFile -LogString $msg
            $Action = "Failed to connect to Exchange Online on import session"
            Write-Log -LogFile $LogFile -LogString $Action
        }
    }
    if ($Connected -eq $true) {
        Write-Log -LogFile $LogFile -LogString "Updating Mailboxes"
        Foreach ($mailbox in $EnabledMailboxes) {
            $UserPrincipalName = "$($mailbox.Alias.ToLower())@$EmailSuffix"
            $i = 0
            $MBX = $null
            Do {
                $MBX = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 30
                $i++
                if ($i -eq 4) {
                    Invoke-ADSync -Cred $Cred -AzureADConnect $AzureADConnect -O365EmailSuffix $O365EmailSuffix
                    $i = 0
                }
            } While (!($MBX) -and $i -lt 5)
            if ($MBX) {
                if ($Mailbox.SharedEquipmentRoom) {
                    $logmsg = "Updating Mailbox: " + $Mailbox.Alias +" "+ $Mailbox.SharedEquipmentRoom +" "+ $Mailbox.Capacity
                    Write-Log -LogFile $LogFile -LogString $LogMsg
                } elseif (!$Mailbox.SharedEquipmentRoom -and !$Mailbox.Capacity) {
                    $logmsg = "Updating Mailbox: " + $Mailbox.Alias
                    Write-Log -LogFile $LogFile -LogString $LogMsg
                }
                Update-UserMailbox -LogFile $LogFile -UserPrincipalName $UserPrincipalName -SharedEquipmentRoom $Mailbox.SharedEquipmentRoom -Capacity $Mailbox.Capacity
            } else {
                $logmsg = "Mailbox: " + $Mailbox.Alias +" not found in AzureAD"
                $Failures += $Mailbox
                Write-Log -LogFile $LogFile -LogString $LogMsg
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
                $UserPrincipalName = "$UserName@$EmailSuffix"
                if ($Dept -eq "IT" -and $HiPriv -eq "Y") {
                    Write-Log -LogFile $LogFile -LogString "Creating Cloud Admin account for $UserName"
                    Write-Log -LogFile $LogFile -LogString " "
                    & $PSScriptRoot\CreateITCloudAdminUser.ps1 -FirstName $FirstName -LastName $LastName -UserName $UserName -EmailSuffix $EmailSuffix -PrivLevel $PrivLevel -Dept $Dept -Company $Company -LogFile $LogFile -Manager $UserName -PasswordLength $PasswordLength
                }
            } catch {
                $e = $_.Exception
                $line = $_.InvocationInfo.ScriptLineNumber
                Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
                Write-Log -LogFile $LogFile -LogString "ERROR processing '$UserName' - Line $line : $($e.Message)" -ForegroundColor Red
                Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
                continue
            }
        }
        Disconnect-ExchangeOnline -Confirm:$false
        Disconnect-MgGraph
    }
    if (Get-PSSession) {
        Write-Log -LogFile $LogFile -LogString "Cleaning up PSSessions"
        Get-PSSession | Remove-PSSession
        $Cred.Password.Dispose()
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    if ($Failures) {
        if ($Failures.Count -gt 0) {
            Foreach ($Failure in $Failures) {
                $UserList | Where-Object {
                    ($_.USERNAME.Trim() -replace '[^A-Za-z0-9.-]','') -ieq $Failure.alias
                } | Export-Csv $FailureFile -NoTypeInformation -Append
            }
        }
    }
}
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString "Processing complete"
Write-Log -LogFile $LogFile -LogString ("=" * 80)
#====================================================================
