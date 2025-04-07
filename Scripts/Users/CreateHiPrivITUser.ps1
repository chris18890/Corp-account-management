[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$O365
    , [Parameter(Mandatory)][string]$EmailSuffix
    , [string]$FirstName,[string]$LastName
    , [string]$UserName,[string]$PasswordLength
    , [string]$Description
    , [string]$Dept,[string]$Company
    , [string]$LogFile,[string]$DCHostName
    , [string]$Manager,[string]$Requester
    , [string]$SMTPServer,[string]$EmailFrom
    , [string]$PrivLevel
)

#====================================================================
# Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$O365 = $O365.ToUpper()
# ADConnect & Exchange settings
if (!$DCHostName) {
    $DCHostName = (Get-ADDomainController).HostName # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
}
$ExServer = "$Domain-EXCH.$DNSSuffix" #Remote Exchange PS session
# Get containing folder for script to locate supporting files
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
# Set variables
$ScriptTitle = "$Domain HiPriv IT User Creation Script"
$OU = "Administration"
$SubOU = "Hi_Priv_Accounts"
$ITAdminGroup = "IT_Admin"
$OUPath = "OU=$SubOU,OU=$OU,$EndPath"
if (!$PasswordLength) {
    $PasswordLength = 4 # Number of characters per password group
}
if (!$PrivLevel) {
    $PrivLevel = READ-HOST 'Enter a Privilege Level for the new account (1-3) - '
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
    "$(Get-Date -Format 'G') $LogString" | Out-File -Filepath $LogFile -Append -Encoding ASCII
    if ($ForegroundColor) {
        Write-Host $LogString -ForegroundColor $ForegroundColor
    } else {
        Write-Host $LogString
    }
}
#====================================================================

#====================================================================
# Generate a random password-legal string
#====================================================================
function Create-Password {
    param([string]$PasswordLength)
    #================================================================
    # Purpose:          Validate password against password policy
    # Assumptions:      Group length has been set and is greater than 3
    # Effects:          Valid password generated
    # Inputs:           $PasswordLength - number of characters for each group
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
# New user email function
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
    }
    # Send email to manager with new user's name & password
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
    } else {
        Write-Log "WARNING: Cannot send email to manager for $UserName, manager field blank..." -ForegroundColor Yellow
    }
}
#====================================================================

if (!$LogFile) {
    $LogFileName = $Domain + "_new_Hi-Priv_user_log-$(Get-Date -Format 'yyyyMMdd')"
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

#====================================================================
# Set user info
#====================================================================
$PrivGroup = "ADM_Role_Level_" + $PrivLevel + "_Admins"
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
    $UserNameAdmin = "admin." + $UserName
} else {
    $UserNameAdmin = "admin." + $UserName
}
if ($UserNameAdmin.Length -gt 20) {
    $UserNameAdmin = $UserNameAdmin.Substring(0,20)
}
$DisplayName = "$LastName, $FirstName (Admin)"
$EmailAddress = "$UserNameAdmin@$EmailSuffix"
$HomeDrive = "H:"
$HomeDir = "\\$Domain\Profiles\$UserNameAdmin"
$UserPrincipalName = "$UserNameAdmin@$EmailSuffix"
#====================================================================
$ExistingUser = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$UserNameAdmin))"
if ($ExistingUser) {
    Write-Log ("-" * 80) -ForegroundColor Red
    Write-Log "ERROR: User '$($UserNameAdmin)' already exists in the $Domain directory. The user object`n is '$($ExistingUser.DistinguishedName)'" -ForegroundColor Red
    Write-Log "Error: Processing '$LastName, $FirstName' ($UserNameAdmin) aborted" -ForegroundColor Red
    Write-Log ("-" * 80) -ForegroundColor Red
    Write-Log ""
    continue # Skip this user
} else {
    try {
        Write-Log ("=" * 80)
        Write-Log "Processing '$DisplayName' ($UserNameAdmin)..."
        Write-Log ("=" * 80)
        #Generate random password
        $UserPassword = Create-Password -PasswordLength $PasswordLength
        #Validate Password against Password Policy
        Validate-Password -Password $UserPassword
        $Params = @{
            Name                    = $UserNameAdmin
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
            Manager                 = $Manager
            SamAccountName          = $UserNameAdmin
            SurName                 = $LastName
            UserPrincipalName       = $UserPrincipalName
        }
        Write-Log "Creating $UserNameAdmin"
        New-ADUser -Type "user" -Server $DCHostName @Params -PassThru
        Set-ADAccountControl -AccountNotDelegated $true -AllowReversiblePasswordEncryption $false -CannotChangePassword $false -DoesNotRequirePreAuth $false -Identity "CN=$UserNameAdmin,$OUPath" -PasswordNeverExpires $false -UseDESKeyOnly $false -Server $DCHostName
        Set-ADObject -Identity "CN=$UserNameAdmin,$OUPath" -protectedFromAccidentalDeletion $True -Server $DCHostName
        Add-GroupMember -Group $ITAdminGroup -Member $UserNameAdmin
        Add-GroupMember -Group $PrivGroup -Member $UserNameAdmin
        Add-GroupMember -Group "Protected Users" -Member $UserNameAdmin
        if ($PrivLevel -ge "3") {
            Write-Log "Creating Domain Admin account for $UserNameAdmin"
            Write-Log ""
            .\CreateDomainAdminUser.ps1 -FirstName $FirstName -LastName $LastName -UserName $UserName -EmailSuffix $EmailSuffix -Description $Description -Dept $Dept -Company $Company -LogFile $LogFile -O365 $O365 -DCHostName $DCHostName -Manager $Manager -Requester $Requester -SMTPServer $SMTPServer -EmailFrom $EmailFrom -PasswordLength $PasswordLength
        }
        if ($O365 -eq "E" -or $O365 -eq "H") {
            # Send email to requester and manager with new user's name & password
            Send-UserEmail -UserName $UserNameAdmin -password $Password -requester $Requester -manager $Manager
        }
        Write-Log ("=" * 80)
        Write-Log "Processing for '$DisplayName' ($UserNameAdmin) complete"
        Write-Log ("=" * 80)
        Write-Log ""
    } catch {
        $e = $_.Exception
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log ("-" * 80) -ForegroundColor Red
        Write-Log "ERROR: Failed during processing of '$DisplayName' ($UserNameAdmin) - Line $Line" -ForegroundColor Red
        Write-Log "$e"
        Write-Log ("-" * 80) -ForegroundColor Red
        Write-Log ("=" * 80)
        Write-Log ""
    }
}
#====================================================================
