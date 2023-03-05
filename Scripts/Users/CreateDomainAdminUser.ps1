[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$O365
    , [Parameter(Mandatory)][string]$EmailSuffix
    , [string]$FirstName,[string]$LastName
    , [string]$UserName,[string]$UserPassword,[string]$PasswordLength
    , [string]$Description
    , [string]$Dept,[string]$Company
    , [string]$O365EmailSuffix
    , [string]$LogFile,[string]$DCHostName
    , [string]$Manager,[string]$Requester
    , [string]$SMTPServer,[string]$EmailFrom
)

#====================================================================
#Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$O365 = $O365.ToUpper()
if ($O365 -eq "H") {
    if (!$O365EmailSuffix) {
        $O365EmailSuffix = READ-HOST 'Enter "onmicrosoft.com" domain - '
        $O365EmailSuffix = "$O365EmailSuffix.onmicrosoft.com"
    }
}
# ADConnect & Exchange settings
if (!$DCHostName) {
    $DCHostName = (Get-ADDomainController).HostName # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
}
$ExServer = "$Domain-EXCH.$DNSSuffix" #Remote Exchange PS session
# Get containing folder for script to locate supporting files
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
# Set variables
$ScriptTitle = "$Domain User Creation Script"
$EnabledMailboxes = @() # Array to Store Completed Mailbox requests for later enumeration
$OU = "IT"
$SubOU = "Hi_Priv_Accounts"
$ITAdminGroup = "IT_Admin"
$OUPath = "OU=$SubOU,OU=$OU,$EndPath"
if (!$PasswordLength) {
    $PasswordLength = 4 # Number of characters per password group
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

#====================================================================
#Set up logging
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
    "$(Get-Date -Format 'G') $LogString" | Out-File -Filepath $LogFile -Append -Encoding ASCII
    if ($ForegroundColor) {
        Write-Host $LogString -ForegroundColor $ForegroundColor
    } else {
        Write-Host $LogString
    }
}
#====================================================================

#====================================================================
#Generate a random password-legal string
#====================================================================
function Create-Password {
    param([string]$PasswordLength)
    #================================================================
    # Purpose:          Validate password against password policy
    # Assumptions:      Group length has been set and is greater than 3
    # Effects:          Valid password generated
    # Inputs:           $Length - number of characters for each group
    # Calls:            Write-Log function
    # Returns:
    # Notes:            There are 4 requirements in the current policy, but this could change in future
    #================================================================
    Write-Log "Generating random password"
    $digits = 48..57
    $lettersLower = 97..122
    $lettersUpper = 65..90
    $passwordLower = get-random -count $PasswordLength -input $lettersLower | % -begin { $aa = $null } -process {$aa += [char]$_} -end {$aa}
    $passwordUpper = get-random -count $PasswordLength -input $lettersUpper | % -begin { $aa = $null } -process {$aa += [char]$_} -end {$aa}
    $passwordDigits = get-random -count $PasswordLength -input $digits | % -begin { $aa = $null } -process {$aa += [char]$_} -end {$aa}
    return $passwordLower + $passwordDigits + $passwordUpper
    Write-Log "Generated random password"
}
#====================================================================

#====================================================================
# Validate password against password policy
#====================================================================
function Validate-Password {
    param([string]$Password)
    #================================================================
    # Purpose:          Validate password against password policy
    # Assumptions:      Password has been generated with enough characters for required groups
    # Effects:          Password should be valid
    # Inputs:           $Password
    # Calls:            Write-Log function
    # Returns:
    # Notes:            There are 4 requirements in the current policy, but this could change in future
    #================================================================
    $TestsPassed = 0
    if ($Password.length -ge 7) {$TestsPassed ++} # Must be >= 7 characters in length
    if ($Password -cmatch "[a-z]") {$TestsPassed ++} # Must contain a lowercase letter
    if ($Password -cmatch "[A-Z]") {$TestsPassed ++} # Must contain an uppercase letter
    if ($Password -match "[0-9]") {$TestsPassed ++} # Must contain a digit
    #if (-Not($Password -notmatch "[a-zA-Z0-9]")) {$TestsPassed ++} # Must contain a special character (not currently required)
    if ($TestsPassed -ge 4) {
        Write-Log "Password validated"
        Write-Log ""
    } else {
        Write-Log ("-" * 80) -ForegroundColor Red
        Write-Log "ERROR: Password '$Password' does not comply with the password policy, script terminating" -ForegroundColor Red
        Write-Log ("-" * 80) -ForegroundColor Red
        exit
    }
}
#====================================================================

#====================================================================
#create mailbox function
#====================================================================
function Create-Mailbox-OnPrem {
    param(
    [string]$UserName
    )
    #================================================================
    # Purpose:          To create an Exchange 2016 Mailbox for a user account
    # Assumptions:      Parameters have been set correctly
    # Effects:          Mailbox should be created for user
    # Inputs:           $UserName - SAM account name of user
    # Calls:            Write-Log function
    # Returns:          $EnabledMailbox
    # Notes:            Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    if (-not $?) {
        $e = $_.Exception
        Write-Log $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log $line -ForegroundColor Red
        $msg = $e.Message
        Write-Log $msg -ForegroundColor Red
        $Action = "ERROR: Error loading Exchange cmdlets - script cannot create Exchange mailbox"
        Write-Log $Action -ForegroundColor Red
    } else {
        #Create Exchange mailbox
        Write-Log "Creating mailbox"
        $alias = $UserName.ToUpper()     #Alias = uppercase UserName
        try {
            $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName"
            Write-log $action
            $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName
            if (-not $?) {
                Write-Log "ERROR: Error creating Exchange mailbox for $UserName - mailbox may not have been created correctly" -ForegroundColor Red
            } else {
                Write-Log "Mailbox created for $UserName successfully"
                $EnabledMailbox = New-Object -Property @{"Alias" = ""} -TypeName PSObject
                $EnabledMailbox.alias = $alias
                Return $EnabledMailbox
            }
        } catch {
            $e = $_.Exception
            Write-Log $e -ForegroundColor Red
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log $line -ForegroundColor Red
            $msg = $e.Message
            Write-Log $msg -ForegroundColor Red
            $Action = "Failed to enable Mailbox or update settings"
            Write-Log $Action -ForegroundColor Red
        }
        Write-Log "End of Mailbox Creation Function"
    }
}
#====================================================================

#====================================================================
#Update mailbox Default Settings
#====================================================================
function Update-Mailbox-OnPrem {
    param(
    [Parameter(Mandatory=$true)] [string]$UserName
    )
    #================================================================
    # Purpose:          Update Mailbox parameters which need to be configured in O365
    # Assumptions:      Parameters have been set correctly
    # Effects:          Mailbox defaults should be assigned to the new mailbox
    # Inputs:           $UserName - SAM account name of user
    # Calls:            Write-Log function
    # Returns:
    # Notes:            Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-Log "Updating Mailbox"
    if (-not $?) {
        Write-Log "ERROR: Error loading Exchange cmdlets - script cannot update Exchange mailbox" -ForegroundColor Red
    } else {
        #Update Exchange mailbox
        $alias = $UserName.ToUpper()     #Alias = uppercase UserName
        $MBX = $null
        try {
            $MBX = Get-Mailbox -Identity $UserName
            $i = 0
            while (!($MBX) -and ($i -le 6)) {
                $MBX = Get-Mailbox -Identity $UserName -erroraction silentlycontinue
                $i++
                Start-Sleep -seconds 10
            }
            if ($MBX) {
                Set-MailboxSpellingConfiguration -Identity $UserName -DictionaryLanguage EnglishUnitedKingdom
                Set-MailboxRegionalConfiguration -Identity $UserName -Language en-GB -DateFormat "dd/MM/yyyy" -TimeFormat "HH:mm" -TimeZone "GMT Standard Time"
                $identityStr = $UserName + ":\Calendar"
                Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Reviewer -DomainController $DCHostName
            }
        } catch {
            $e = $_.Exception
            Write-Log $e -ForegroundColor Red
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log $line -ForegroundColor Red
            $msg = $e.Message
            Write-Log $msg -ForegroundColor Red
            $Action = "Failed to Complete Mailbox Update"
            Write-Log $Action -ForegroundColor Red
        }
    }
    Write-Log "End of Mailbox Update Function"
}
#====================================================================

#====================================================================
#create mailbox function
#====================================================================
function Create-Mailbox-Hybrid {
    param(
    [string]$UserName
    )
    #================================================================
    # Purpose:          To create an Exchange online Mailbox for a user account
    # Assumptions:      Parameters have been set correctly
    # Effects:          Mailbox should be created for user
    # Inputs:           $UserName - SAM account name of user
    # Calls:            Write-Log function
    # Returns:          $EnabledMailbox
    # Notes:            Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    if (-not $?) {
        $e = $_.Exception
        Write-Log $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log $line -ForegroundColor Red
        $msg = $e.Message
        Write-Log $msg -ForegroundColor Red
        $Action = "ERROR: Error loading Exchange cmdlets - script cannot create Exchange mailbox"
        Write-Log $Action -ForegroundColor Red
    } else {
        #Create Exchange mailbox
        Write-Log "Creating mailbox"
        $alias = $UserName.ToUpper()     #Alias = uppercase UserName
        try {
            $action = "Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix"
            Write-log $action
            $NewMailbox = Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix"
            if (-not $?) {
                Write-Log "ERROR: Error creating Exchange mailbox for $UserName - mailbox may not have been created correctly" -ForegroundColor Red
            } else {
                Write-Log "Mailbox created for $UserName successfully"
                $EnabledMailbox = New-Object -Property @{"Alias" = ""} -TypeName PSObject
                $EnabledMailbox.alias = $alias
                Return $EnabledMailbox
            }
        } catch {
            $e = $_.Exception
            Write-Log $e -ForegroundColor Red
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log $line -ForegroundColor Red
            $msg = $e.Message
            Write-Log $msg -ForegroundColor Red
            $Action = "Failed to enable Mailbox or update settings"
            Write-Log $Action -ForegroundColor Red
        }
        Write-Log "End of Mailbox Creation Function"
    }
}
#====================================================================

#====================================================================
#Update mailbox Default Settings
#====================================================================
function Update-Mailbox-Hybrid {
    param(
    [Parameter(Mandatory=$true)] [string]$UserName
    )
    #================================================================
    # Purpose:          Update Mailbox parameters which need to be configured in O365
    # Assumptions:      Parameters have been set correctly
    # Effects:          Mailbox defaults should be assigned to the new mailbox
    # Inputs:           $UserName - SAM account name of user
    # Calls:            Write-Log function
    # Returns:
    # Notes:            Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-Log "Updating Mailbox"
    if (-not $?) {
        Write-Log "ERROR: Error loading Exchange cmdlets - script cannot update Exchange mailbox" -ForegroundColor Red
    } else {
        #Update Exchange mailbox
        $alias = $UserName.ToUpper()     #Alias = uppercase UserName
        $MBX = $null
        try {
            $MBX = Get-Mailbox -Identity $UserName
            $i = 0
            while (!($MBX) -and ($i -le 6)) {
                $MBX = Get-Mailbox -Identity $UserName -erroraction silentlycontinue
                $i++
                Start-Sleep -seconds 10
            }
            if ($MBX) {
                Set-MailboxSpellingConfiguration -Identity $UserName -DictionaryLanguage EnglishUnitedKingdom
                Set-MailboxRegionalConfiguration -Identity $UserName -Language en-GB -DateFormat "dd/MM/yyyy" -TimeFormat "HH:mm" -TimeZone "GMT Standard Time"
                $identityStr = $UserName + ":\Calendar"
                Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Reviewer #-DomainController $DCHostName
            }
        } catch {
            $e = $_.Exception
            Write-Log $e -ForegroundColor Red
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log $line -ForegroundColor Red
            $msg = $e.Message
            Write-Log $msg -ForegroundColor Red
            $Action = "Failed to Complete Mailbox Update"
            Write-Log $Action -ForegroundColor Red
        }
    }
    Write-Log "End of Mailbox Update Function"
}
#====================================================================

#====================================================================
function Test-Cred {
    [CmdletBinding()]
    [OutputType([String])]
    Param (
        [Parameter(
            Mandatory = $false,
            ValueFromPipeLine = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias(
            'PSCredential'
        )]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credentials
    )
    #================================================================
    # Purpose:          Test Provided Credentials Prior to Use
    # Assumptions:      
    # Effects:
    # Inputs:           $Cred
    # $LogString:       
    # Calls:
    # Returns:         Authenticated or Unauthenticated
    # Notes:
    #================================================================
    $Domain = $null
    $Root = $null
    $Username = $null
    $Password = $null
    if ($Credentials -eq $null) {
        try {
            $Credentials = Get-Credential "domain\$env:username" -ErrorAction Stop
        } catch {
            $ErrorMsg = $_.Exception.Message
            Write-Warning "Failed to validate credentials: $ErrorMsg "
            Pause
            Break
        }
    }
    # Checking module
    try {
        # Split username and password
        $Username = $credentials.username
        $Password = $credentials.GetNetworkCredential().password
        # Get Domain
        $Root = "LDAP://" + ([ADSI]'').distinguishedName
        $Domain = New-Object System.DirectoryServices.DirectoryEntry($Root,$UserName,$Password)
    } catch {
        $_.Exception.Message
        Continue
    }
    if (!$domain) {
        Write-Warning "Something went wrong"
    } else {
        if ($domain.name -ne $null) {
            return "Authenticated"
        } else {
            return "Not authenticated"
        }
    }
}
#====================================================================

#====================================================================
#group addition function
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
    $Error.Clear()
    $checkGroup = Get-ADGroup -LDAPFilter "(SAMAccountName=$Group)"
    if ($checkGroup -ne $null) {
        Write-Log "Adding $Member to $Group"
        try {
            Add-ADGroupMember -Identity $Group -Members $Member -Server $DCHostName
            Write-Log "Added $Member to $Group"
        }
        catch [Microsoft.ActiveDirectory.Management.ADException] {
            switch ($Error[0].Exception.ErrorCode) {
                1378 { # 'The specified object is already a member of the group'
                    Write-Log "'$Member' is already a member of group '$Group'" -ForegroundColor Yellow
                }
                default {
                    Write-Log "ERROR: An unexpected error occurred while attempting to add user '$Member' to a group:`n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                }
            }
        }
    } else {
        Write-Log "$Group does not exist" -ForegroundColor Red
    }
}
#====================================================================

#====================================================================
#new user email function
#====================================================================
function Send-UserEmail {
    param([string]$UserName,[string]$Password,[string]$Requester,[string]$Manager)
    #================================================================
    # Purpose:          To send an email to the requester and/or manager of the new account
    # Assumptions:      Parameters have been set correctly
    # Effects:          Email will be sent to the Requester if field is not blank
    #                   If the manager is different from the requestor, an email
    #                   will also be sent to the manager provided the field is not blank
    # Inputs:           $UserName - SAM account name of user
    #                   $Password - Password for the user account
    #                   $Requester - Person who requested the account, from the CSV
    #                   $Manager - User's manager, as set on the org tab of account properties
    # Calls:            Write-Log function
    # Returns:
    # Notes:
    #================================================================
    # Send email to requester with new user's name & password
    if ($Requester) {
        $CheckRequester = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$Requester))"
        if ($CheckRequester) {
            Write-Log "Sending email to requester ($Requester) for $UserName..."
            $RequesterEmail = Get-ADUser $Requester -Properties mail | Select-Object -ExpandProperty mail
            $UserEmail = Get-ADUser $UserName -Properties mail | Select-Object -ExpandProperty mail
            $DisplayName = Get-ADUser $UserName -Properties DisplayName | Select-Object -ExpandProperty DisplayName
            $Splat = @{
                To          = $RequesterEmail
                From        = "$ScriptTitle <$EmailFrom>"
                Body        = "New User Created`n`nUserName is $UserName,`nEmail address is $UserEmail,`n`nPassword is $Password.`n`n`nPlease do not reply to this email, it has been sent from an unmonitored address."
                Subject     = "New User Created - $DisplayName"
                SmtpServer  = $SMTPServer
                Priority    = "High"
            }
            Send-MailMessage @Splat
        } else {
            Write-Log "WARNING: Cannot send email to requester for $UserName, requester field incorrect..." -ForegroundColor Yellow
        }
    }# Send email to manager with new user's name & password
    if ($Manager) {
        $CheckManager = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$Manager))"
        if ($CheckManager) {
            if ($Manager -ne $Requester) { # check to see if manager is the same as requester, only send email if they're different
                Write-Log "Sending email to manager ($Manager) for $UserName..."
                $ManagerEmail = Get-ADUser $Manager -Properties mail | Select-Object -ExpandProperty mail
                $UserEmail = Get-ADUser $UserName -Properties mail | Select-Object -ExpandProperty mail
                $DisplayName = Get-ADUser $UserName -Properties DisplayName | Select-Object -ExpandProperty DisplayName
                $Splat = @{
                    To          = $ManagerEmail
                    From        = "$ScriptTitle <$EmailFrom>"
                    Body        = "New User Created`n`nUserName is $UserName,`nEmail address is $UserEmail,`n`nPassword is $Password.`n`n`nPlease do not reply to this email, it has been sent from an unmonitored address."
                    Subject     = "New User Created - $DisplayName"
                    SmtpServer  = $SMTPServer
                    Priority    = "High"
                }
                Send-MailMessage @Splat
            }
        } else {
            Write-Log "WARNING: Cannot send email to manager for $UserName, manager field incorrect..." -ForegroundColor Yellow
        }
    }
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
    #Generate Password if one hasn't been passed as a param
    if (!$UserPassword) {
        $UserPassword = Create-Password -PasswordLength $PasswordLength
        #Validate Password against Password Policy
        Validate-Password -Password $UserPassword
    }
    if ($O365 -eq "E" -or $O365 -eq "H") {
        # Get user credentials for server connectivity (Non-MFA)
        try {
            $Cred = Get-Credential -ErrorAction Stop -Message "Admin credentials for remote sessions:"
        } catch {
            $ErrorMsg = $_.Exception.Message
            Write-Log "Failed to validate credentials: $ErrorMsg "
            Pause
            Break
        }
        $CredCheck = $Cred | Test-Cred
        if ($CredCheck -ne "Authenticated") {
            Write-Log "Credential validation failed - Script Terminating"
            pause
            Exit
        }
        #Connect to remote Exchange PowerShell
        Write-Log "Connecting to remote Exchange PowerShell session... "
        try {
            $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ExServer/PowerShell" -Name ExSession -Credential $cred -ErrorAction Stop -authentication Kerberos
            $ExConnected = $true
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
            Write-log "Exchange session not connected Stopping Script"
            Exit
        }
    }
}
#====================================================================

#====================================================================
# Set user info
#====================================================================
$Membership = "$ITAdminGroup"
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
    $UserName = $UserName.Trim() -replace '[^A-Za-z0-9.]', [String]::Empty # Strip out illegal characters from User ID
    $UserNameDomainAdmin = "da." + $UserName
} else {
    $UserNameDomainAdmin = "da." + $UserName
}
if ($UserNameDomainAdmin.Length -gt 20) {
    $UserNameAdmin = $UserNameDomainAdmin.Substring(0,20)
}
$DisplayName = "$LastName, $FirstName (Domain Admin)"
$EmailAddress = "$UserNameDomainAdmin@$EmailSuffix"
$HomeDrive = "H:"
$HomeDir = "\\$Domain\Profiles\$UserNameDomainAdmin"
$UserPrincipalName = "$UserNameDomainAdmin@$EmailSuffix"
#====================================================================
$ExistingUser = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$UserNameDomainAdmin))"
if ($ExistingUser) {
    Write-Log ("-" * 80) -ForegroundColor Red
    Write-Log "ERROR: User '$($UserNameDomainAdmin)' already exists in the $Domain directory. The user object`n is '$($ExistingUser.DistinguishedName)'" -ForegroundColor Red
    Write-Log "Error: Processing '$LastName, $FirstName' ($UserNameDomainAdmin) aborted" -ForegroundColor Red
    Write-Log ("-" * 80) -ForegroundColor Red
    Write-Log ""
    continue # Skip this user
} else {
    try {
        Write-Log ("=" * 80)
        Write-Log "Processing '$DisplayName' ($UserNameDomainAdmin)..."
        Write-Log ("=" * 80)
        $Params = @{
            Name                    = $UserNameDomainAdmin
            AccountPassword         = ConvertTo-SecureString -AsPlainText $UserPassword -Force
            ChangePasswordAtLogon   = $false
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
        Write-Log "Creating $UserNameDomainAdmin"
        New-ADUser -Type "user" -Server $DCHostName @Params -PassThru
        Set-ADAccountControl -AccountNotDelegated $false -AllowReversiblePasswordEncryption $false -CannotChangePassword $false -DoesNotRequirePreAuth $false -Identity "CN=$UserNameDomainAdmin,$OUPath" -PasswordNeverExpires $false -UseDESKeyOnly $false -Server $DCHostName
        Add-GroupMember -Group $Membership -Member $UserNameDomainAdmin
        Add-GroupMember -Group $PrivGroup -Member $UserNameDomainAdmin
        Add-GroupMember -Group "Protected Users" -Member $UserNameDomainAdmin
        if ($O365 -eq "E") {
            Write-Log "Exchange mailbox for $UserNameDomainAdmin will be created in Exchange OnPrem"
            Write-Log "Calling Create-Mailbox-OnPrem function with the following parameters:"
            Write-Log "UserName: $UserNameDomainAdmin"
            $enabledMailboxes += Create-Mailbox-OnPrem -UserName $UserNameDomainAdmin
            Write-log "Updating Mailboxes"
            foreach ($mailbox in $EnabledMailboxes) {
                $i = 0
                $MBX = $null
                Do {
                    $MBX = Get-Mailbox $mailbox.alias -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 30
                    $i++
                } While (!($MBX) -and $i -lt 5)
                if ($MBX) {
                    $logmsg = "Updating Mailbox:" + $Mailbox.Alias
                    Write-log $logMsg
                    Update-Mailbox-OnPrem $mailbox.Alias
                } else {
                    $logmsg = "Mailbox:" + $Mailbox.Alias +" not found in AD"
                    Write-log $logMsg
                }
            }
            # Send email to requester and manager with new user's name & password
            Send-UserEmail -UserName $UserNameDomainAdmin -password $Password -requester $Requester -manager $Manager
        }
        if ($O365 -eq "H") {
            Write-Log "Exchange mailbox for $UserNameDomainAdmin will be created in Exchange Online"
            Write-Log "Calling Create-Mailbox-Hybrid function with the following parameters:"
            Write-Log "UserName: $UserNameDomainAdmin"
            $enabledMailboxes += Create-Mailbox-Hybrid -UserName $UserNameDomainAdmin
            # Send email to requester and manager with new user's name & password
            Send-UserEmail -UserName $UserNameDomainAdmin -password $Password -requester $Requester -manager $Manager
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
