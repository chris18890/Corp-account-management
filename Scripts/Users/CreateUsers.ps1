[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$O365
    , [Parameter(Mandatory)][string]$EmailSuffix
    , [string]$O365EmailSuffix
)

#====================================================================
#Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$UsersOU = "Staff"
$O365 = $O365.ToUpper()
if ($O365 -eq "E" -or $O365 -eq "H") {
    if (!$O365EmailSuffix) {
        $O365EmailSuffix = READ-HOST 'Enter "onmicrosoft.com" domain - '
    }
    $O365EmailSuffix = "$O365EmailSuffix.onmicrosoft.com"
}
# Group settings
$GroupsOU = "Groups"
$GroupCategory = "Security"
$GroupScope = "Universal"
$SharedGroupsOU = "OU=Shared_Mailbox_Access,OU=$GroupsOU"
$EquipmentGroupsOU = "OU=Equipment_Mailbox_Access,OU=$GroupsOU"
$RoomGroupsOU = "OU=Room_Mailbox_Access,OU=$GroupsOU"
$SharedAccountsOU = "OU=Shared_Mailbox_Accounts"
$EquipmentAccountsOU = "OU=Equipment_Mailbox_Accounts"
$RoomAccountsOU = "OU=Room_Mailbox_Accounts"
$O365LicenseGroup = "UG_Office365"
# ADConnect & Exchange settings
$DCHostName = (Get-ADDomainController).HostName # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ExServer = "$Domain-EXCH.$DNSSuffix" #Remote Exchange PS session
$AzureADConnect = "$Domain-RTR.$DNSSuffix"
# Get containing folder for script to locate supporting files
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
# Set variables
$ScriptTitle = "$Domain User Creation Script"
$SMTPServer = "$Domain-EXCH.$DNSSuffix" # SMTP server used for email notifications
$EmailFrom = "noreply@$EmailSuffix" # From address
$PasswordLength = 4 # Number of characters per password group
$EnabledMailboxes = @() # Array to Store Completed Mailbox requests for later enumeration
$Roles = @("Company Administrator")
$Level1Roles = @("Helpdesk Administrator", "Service support administrator", "Global Reader")
$Level2Roles = @("User Administrator", "Groups administrator", "Authentication administrator", "License Administrator")
$Level3Roles = @("Exchange Administrator", "Teams Administrator", "Sharepoint Administrator", "Privileged authentication administrator", "Privileged role administrator")
# File locations
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Log "Creating log folder"
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
    [string]$UserName,[string]$realname
    ,[string]$SharedEquipmentRoom,[string]$Capacity
    )
    #================================================================
    # Purpose:          To create an Exchange 2019 Mailbox for a user account
    # Assumptions:      Parameters have been set correctly
    # Effects:          Mailbox should be created for user
    # Inputs:           $UserName - SAM account name of user
    #                   $realname - Real Name to set as Primary SMTP address, read from CSV
    #                   $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #                   $Capacity - If mailbox is a room account, set the capacity
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
            switch ($SharedEquipmentRoom) {
                "S" {
                    if ($realname) {
                        $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -shared"
                        Write-Log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -shared
                    } else {
                        $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -shared"
                        Write-Log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -shared
                    }
                }
                "E" {
                    if ($realname) {
                        $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -equipment"
                        Write-Log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -equipment
                    } else {
                        $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -equipment"
                        Write-Log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -equipment
                    }
                }
                "R" {
                    if ($realname) {
                        $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -room"
                        Write-Log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -room
                    } else {
                        $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -room"
                        Write-Log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -room
                    }
                }
                default {
                    if ($realname) {
                        $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName"
                        Write-Log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName
                    } else {
                        $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName"
                        Write-Log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName
                    }
                }
            }
            if (-not $?) {
                Write-Log "ERROR: Error creating Exchange mailbox for $UserName - mailbox may not have been created correctly" -ForegroundColor Red
            } else {
                Write-Log "Mailbox created for $UserName successfully"
                $EnabledMailbox = New-Object -Property @{"Alias" = ""; "SharedEquipmentRoom" = ""; "Capacity" = ""} -TypeName PSObject
                $EnabledMailbox.alias = $alias
                $EnabledMailbox.SharedEquipmentRoom = $SharedEquipmentRoom
                $EnabledMailbox.Capacity = $Capacity
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
    , [Parameter(Mandatory=$false)] [string]$SharedEquipmentRoom = ""
    , [Parameter(Mandatory=$false)] [string]$Capacity = ""
    )
    #================================================================
    # Purpose:          Update Mailbox parameters which need to be configured in O365
    # Assumptions:      Parameters have been set correctly
    # Effects:          Mailbox defaults should be assigned to the new mailbox
    # Inputs:           $UserName - SAM account name of user
    #                   $realname - Real Name to set as Primary SMTP address, read from CSV
    #                   $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #                   $Capacity - If mailbox is a room account, set the capacity
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
                #Set shared account settings
                switch ($SharedEquipmentRoom) {
                    "S" {
                        Write-Log "Updating Shared Mailbox $UserName : Adding Permissions"
                        $GroupName = "sh_$UserName"
                        Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false -DomainController $DCHostName
                        Add-ADPermission -Identity $UserName -User $GroupName -ExtendedRights "Send As" -confirm:$false -DomainController $DCHostName
                        Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                    }
                    "E" {
                        Write-Log "Updating Equipment Mailbox $UserName : Adding Permissions"
                        #Set Default calendar permissions to Author
                        $MBXType = (Get-Mailbox -Identity $UserName).RecipientTypeDetails
                        $x = 0
                        While ($MBXType -ne "EquipmentMailbox" -and $x -lt 6) {
                            Start-Sleep -Seconds 10
                            $MBXType = (Get-Mailbox -Identity $UserName).RecipientTypeDetails
                            $x++
                        }
                        $identityStr = $UserName + ":\Calendar"
                        Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -confirm:$false -DomainController $DCHostName
                        #Set calendar resource attendant to auto-accept
                        Write-Log "Updating Equipment Mailbox $UserName : Updating Calendar Processing"
                        Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false -DomainController $DCHostName
                        $GroupName = "eq_$UserName"
                        Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false -DomainController $DCHostName
                        Add-ADPermission -Identity $UserName -User $GroupName -ExtendedRights "Send As" -confirm:$false -DomainController $DCHostName
                        Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                    }
                    "R" {
                        Write-Log "Updating Room Mailbox $UserName : Adding Permissions"
                        #Set Default calendar permissions to Author
                        $MBXType = (Get-Mailbox -Identity $UserName).RecipientTypeDetails
                        $x = 0
                        While ($MBXType -ne "RoomMailbox" -and $x -lt 6) {
                            Start-Sleep -Seconds 10
                            $MBXType = (Get-Mailbox -Identity $UserName).RecipientTypeDetails
                            $x++
                        }
                        $identityStr = $UserName + ":\Calendar"
                        Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -confirm:$false -DomainController $DCHostName
                        #Set calendar resource attendant to auto-accept
                        Write-Log "Updating Room Mailbox $UserName : Updating Calendar Processing"
                        Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false -DomainController $DCHostName
                        if ($Capacity -ne "N") {
                            Set-Mailbox -Identity $UserName -ResourceCapacity $Capacity -DomainController $DCHostName
                        }
                        $GroupName = "ro_$UserName"
                        Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false -DomainController $DCHostName
                        Add-ADPermission -Identity $UserName -User $GroupName -ExtendedRights "Send As" -confirm:$false -DomainController $DCHostName
                        Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                    }
                }
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
    [string]$UserName,[string]$realname
    ,[string]$SharedEquipmentRoom,[string]$Capacity
    )
    #================================================================
    # Purpose:          To create an Exchange Online Mailbox for a user account
    # Assumptions:      Parameters have been set correctly
    # Effects:          Mailbox should be created for user
    # Inputs:           $UserName - SAM account name of user
    #                   $realname - Real Name to set as Primary SMTP address, read from CSV
    #                   $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #                   $Capacity - If mailbox is a room account, set the capacity
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
            Switch ($SharedEquipmentRoom) {
                "S" {
                    if ($realname) {
                        $action = "Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -shared"
                        Write-Log $action
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -shared
                    } else {
                        $action = "Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -shared"
                        Write-Log $action
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -shared
                    }
                }
                "E" {
                    if ($realname) {
                        $action = "Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -equipment"
                        Write-Log $action
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -equipment
                    } else {
                        $action = "Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -equipment"
                        Write-Log $action
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -equipment
                    }
                }
                "R" {
                    if ($realname) {
                        $action = "Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -room"
                        Write-Log $action
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -room
                    } else {
                        $action = "Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -room"
                        Write-Log $action
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -room
                    }
                }
                default {
                    if ($realname) {
                        $action = "Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix"
                        Write-Log $action
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix"
                    } else {
                        $action = "Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix"
                        Write-Log $action
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix"
                    }
                }
            }
            if (-not $?) {
                Write-Log "ERROR: Error creating Exchange mailbox for $UserName - mailbox may not have been created correctly" -ForegroundColor Red
            } else {
                Write-Log "Mailbox created for $UserName successfully"
                $EnabledMailbox = New-Object -Property @{"Alias" = ""; "SharedEquipmentRoom" = ""; "Capacity" = ""} -TypeName PSObject
                $EnabledMailbox.alias = $alias
                $EnabledMailbox.SharedEquipmentRoom = $SharedEquipmentRoom
                $EnabledMailbox.Capacity = $Capacity
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
    , [Parameter(Mandatory=$false)] [string]$SharedEquipmentRoom = ""
    , [Parameter(Mandatory=$false)] [string]$Capacity = ""
    )
    #================================================================
    # Purpose:          Update Mailbox parameters which need to be configured in O365
    # Assumptions:      Parameters have been set correctly
    # Effects:          Mailbox defaults should be assigned to the new mailbox
    # Inputs:           $UserName - SAM account name of user
    #                   $realname - Real Name to set as Primary SMTP address, read from CSV
    #                   $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #                   $Capacity - If mailbox is a room account, set the capacity
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
                #Set shared account settings
                Switch ($SharedEquipmentRoom) {
                    "S" {
                        if ($MBX.RecipientTypeDetails -ne "SharedMailbox") {
                            Write-Log "Converting $UserName Mailbox to shared type"
                            Set-Mailbox -Identity $UserName -type:shared
                        }
                        Write-Log "Updating Shared Mailbox $UserName : Adding Permissions"
                        $GroupName = "sh_$UserName"
                        Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false #-DomainController $DCHostName
                        Add-RecipientPermission -Identity $UserName -trustee $GroupName -Accessrights "SendAs" -confirm:$false #-DomainController $DCHostName
                        Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                    }
                    "E" {
                        if ($MBX.RecipientTypeDetails -ne "EquipmentMailbox") {
                            Write-Log "Converting $UserName Mailbox to equipment type"
                            Set-Mailbox -Identity $UserName -type:equipment
                        }
                        Write-Log "Updating Equipment Mailbox $UserName : Adding Permissions"
                        #Set Default calendar permissions to Author
                        $MBXType = (Get-Mailbox -Identity $UserName).RecipientTypeDetails
                        $x = 0
                        While ($MBXType -ne "EquipmentMailbox" -and $x -lt 6) {
                            Start-Sleep -Seconds 10
                            $MBXType = (Get-Mailbox -Identity $UserName).RecipientTypeDetails
                            $x++
                        }
                        $identityStr = $UserName + ":\Calendar"
                        Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -confirm:$false #-DomainController $DCHostName
                        #Set calendar resource attendant to auto-accept
                        Write-Log "Updating Equipment Mailbox $UserName : Updating Calendar Processing"
                        Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false #-DomainController $DCHostName
                        $GroupName = "eq_$UserName"
                        Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false #-DomainController $DCHostName
                        Add-RecipientPermission -Identity $UserName -Trustee $GroupName -Accessrights "SendAs" -confirm:$false #-DomainController $DCHostName
                        Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                    }
                    "R" {
                        if ($MBX.RecipientTypeDetails -ne "RoomMailbox") {
                            Write-Log "Converting $UserName Mailbox to room type"
                            Set-Mailbox -Identity $UserName -type:room
                        }
                        Write-Log "Updating Room Mailbox $UserName : Adding Permissions"
                        #Set Default calendar permissions to Author
                        $MBXType = (Get-Mailbox -Identity $UserName).RecipientTypeDetails
                        $x = 0
                        While ($MBXType -ne "RoomMailbox" -and $x -lt 6) {
                            Start-Sleep -Seconds 10
                            $MBXType = (Get-Mailbox -Identity $UserName).RecipientTypeDetails
                            $x++
                        }
                        $identityStr = $UserName + ":\Calendar"
                        Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -confirm:$false #-DomainController $DCHostName
                        #Set calendar resource attendant to auto-accept
                        Write-Log "Updating Room Mailbox $UserName : Updating Calendar Processing"
                        Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false #-DomainController $DCHostName
                        if ($Capacity -ne "N") {
                            Set-Mailbox -Identity $UserName -ResourceCapacity $Capacity #-DomainController $DCHostName
                        }
                        $GroupName = "ro_$UserName"
                        Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false #-DomainController $DCHostName
                        Add-RecipientPermission -Identity $UserName -Trustee $GroupName -Accessrights "SendAs" -confirm:$false #-DomainController $DCHostName
                        Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                    }
                }
                #Write-Log "Enabling Litigation hold on Mailbox $UserName"
                #Set-Mailbox -Identity $UserName -LitigationHoldEnabled $true -LitigationHoldDuration 365
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
#AD Sync
#====================================================================
function Force-ADSync {
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
#group creation function
#====================================================================
function Create-ADGroup {
    [CmdletBinding()]
    param(
        [string]$GroupName,[String]$Path,[String]$GroupDescription
    )
    $Error.Clear()
    try {
        New-ADGroup -GroupCategory $GroupCategory -GroupScope $GroupScope -Name $GroupName -OtherAttributes:@{mail="$GroupName@$EmailSuffix"} -Path $Path -SamAccountName $GroupName -Server $DCHostName -Description $GroupDescription
    }
    catch [Microsoft.ActiveDirectory.Management.ADException] {
        Write-Log "'$GroupName' already exists" -ForegroundColor Green
    }
    if ($O365 -eq "E" -or $O365 -eq "H") {
        Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
        Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
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

Write-Log ("=" * 80)
Write-Log "Log file is '$LogFile'"
Write-Log "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log "DC being used is '$DCHostName'"
Write-Log "Script path is '$ScriptPath'"
Write-Log "$ScriptTitle"
Write-Log "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log ("=" * 80)
Write-Log ""
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
        Write-Log "Exchange session not connected Stopping Script"
        Exit
    }
}
#====================================================================

#====================================================================
#Loop through CSV & create users
#====================================================================
$LIST = @(IMPORT-CSV "users.csv")
$CreatedUsers = @()
foreach ($USER in $LIST) {
    $Membership = "$UsersOU"
    $FirstName = $USER.FIRSTNAME
    $FirstName = $FirstName.Trim() -replace '[?@\\+]', [String]::Empty # Strip out illegal characters from First name for Office 365 compliance. Note that \ is escaped to \\
    $LastName = $USER.LASTNAME
    $LastName = $LastName.Trim() -replace '[?@\\+]', [String]::Empty # Strip out illegal characters from LastName for Office 365 compliance. Note that \ is escaped to \\
    $UserName = $USER.USERNAME
    $UserName = $UserName.Trim() -replace '[^A-Za-z0-9.]', [String]::Empty # Strip out illegal characters from User ID
    $Description = $USER.Description
    $Company = $USER.COMPANY
    $Dept = $USER.DEPT
    $HiPriv = $USER.HIPRIV.ToUpper()
    $PrivLevel = $USER.PrivLevel
    $SharedEquipmentRoom = $USER.'S/E/R'.ToUpper()
    $AdminID = $USER.AdminID.ToLower()
    $Managed = $USER.Managed.ToUpper()
    $Capacity = $USER.Cap
    $Manager = $USER.MANAGER.ToLower()
    $Requester = $User.Requester.ToLower()
    $Phone = $User.PHONE
    switch ($SharedEquipmentRoom) {
        "S" {
            $DisplayName = $FirstName
            $LastName = "Shared"
            $OUPath = "$SharedAccountsOU,OU=IT,$EndPath"
            $Enabled = $false
        }
        "E" {
            $DisplayName = $FirstName
            $LastName = "Equipment"
            $OUPath = "$EquipmentAccountsOU,OU=IT,$EndPath"
            $Enabled = $false
        }
        "R" {
            $DisplayName = $FirstName
            $LastName = "Room"
            $OUPath = "$RoomAccountsOU,OU=IT,$EndPath"
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
    $HomeDir = "\\$Domain\Profiles\$UserName"
    $UserPrincipalName = "$UserName@$EmailSuffix"
    $ExistingUser = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$UserName))"
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
            #Generate random password
            $UserPassword = Create-Password -PasswordLength $PasswordLength
            #Validate Password against Password Policy
            Validate-Password -Password $UserPassword
            $Params = @{
                Name                    = $UserName
                AccountPassword         = ConvertTo-SecureString -AsPlainText $UserPassword -Force
                ChangePasswordAtLogon   = $true
                Description             = $Description
                Company                 = $Company
                Department              = $Dept
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
            New-ADUser -Type "user" -Server $DCHostName @Params -PassThru
            Write-Log "Created $UserName"
            Set-ADAccountControl -AccountNotDelegated $false -AllowReversiblePasswordEncryption $false -CannotChangePassword $false -DoesNotRequirePreAuth $false -Identity "CN=$UserName,$OUPath" -PasswordNeverExpires $false -UseDESKeyOnly $false -Server $DCHostName
            Add-GroupMember -Group $Membership -Member $UserName
            $objNewUserDE = [ADSI]"LDAP://CN=$UserName,$OUPath"
            if ($Dept) {
                Add-GroupMember -Group $Dept -Member $UserName
            }
            switch ($SharedEquipmentRoom) {
                "S" {
                    # create management group for shared account
                    $GroupName = "sh_$UserName"
                    Write-Log "Creating group $GroupName for shared account management"
                    Create-ADGroup -GroupName $GroupName -Path "$SharedGroupsOU,$EndPath" -GroupDescription "Group to grant access access to $UserName"
                    Write-Log "Group $GroupName created in location $SharedGroupsOU,$EndPath"
                    if ($AdminID) {
                        $CheckAdminID = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$AdminID))"
                        if ($CheckAdminID) {
                            Add-GroupMember -Group $GroupName -Member $AdminID
                        }
                    }
                }
                "E" {
                    # create management group for equipment account
                    $GroupName = "eq_$UserName"
                    Write-Log "Creating group $GroupName for equipment account management"
                    Create-ADGroup -GroupName $GroupName -Path "$EquipmentGroupsOU,$EndPath" -GroupDescription "Group to grant access access to $UserName"
                    Write-Log "Group $GroupName created in location $EquipmentGroupsOU,$EndPath"
                    if ($AdminID) {
                        if ($Managed -eq "M") {
                            $Assistant = $AdminID + " (M)"
                        } else {
                            $Assistant = $AdminID
                        }
                        Write-Log "Setting Assistant for Equipment Account to '$Assistant'..."
                        $objNewUserDE.PSBase.InvokeSet("msExchAssistantName",$Assistant)
                        $CheckAdminID = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$AdminID))"
                        if ($CheckAdminID) {
                            Add-GroupMember -Group $GroupName -Member $AdminID
                        }
                    }
                    $objNewUserDE.SetInfo() # Commit changes
                }
                "R" {
                    # create management group for room account
                    $GroupName = "ro_$UserName"
                    Write-Log "Creating group $GroupName for room account management"
                    Create-ADGroup -GroupName $GroupName -Path "$RoomGroupsOU,$EndPath" -GroupDescription "Group to grant access access to $UserName"
                    Write-Log "Group $GroupName created in location $RoomGroupsOU,$EndPath"
                    if ($Capacity) {
                        Write-Log "Setting Title for Room Account to 'Cap: $Capacity'..."
                        $objNewUserDE.PSBase.InvokeSet("title","Cap: $Capacity")
                    }
                    if ($AdminID) {
                        if ($Managed -eq "M") {
                            $Assistant = $AdminID + " (M)"
                        } else {
                            $Assistant = $AdminID
                        }
                        Write-Log "Setting Assistant for Room Account to '$Assistant'..."
                        $objNewUserDE.PSBase.InvokeSet("msExchAssistantName",$Assistant)
                        $CheckAdminID = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$AdminID))"
                        if ($CheckAdminID) {
                            Add-GroupMember -Group $GroupName -Member $AdminID
                        }
                    }
                    $objNewUserDE.SetInfo() # Commit changes
                }
                default {
                    if ($O365LicenseGroup) {
                        Add-GroupMember -Group $O365LicenseGroup -Member $UserName
                    }
                    # Set manager
                    if ($Manager) {
                        $CheckManager = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$Manager))"
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
                        .\CreateHiPrivITUser.ps1 -FirstName $FirstName -LastName $LastName -UserName $UserName -UserPassword $UserPassword -EmailSuffix $EmailSuffix -PrivLevel $PrivLevel -Description $Description -Dept $Dept -Company $Company -LogFile $LogFile -O365 $O365 -O365EmailSuffix $O365EmailSuffix -DCHostName $DCHostName -Manager $Manager -Requester $Requester -SMTPServer $SMTPServer -EmailFrom $EmailFrom -PasswordLength $PasswordLength
                    }
                }
            }
            if ($O365 -eq "E") {
                Write-Log "Exchange mailbox for $UserName will be created in Exchange OnPrem"
                Write-Log "Calling Create-Mailbox-OnPrem function with the following parameters:"
                Write-Log "UserName: $UserName, realname: $RealName, SharedEquipmentRoom: $SharedEquipmentRoom, Capacity: $Capacity"
                $enabledMailboxes += Create-Mailbox-OnPrem -UserName $UserName -realname $RealName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
                # Send email to requester and manager with new user's name & password
                if ($SharedEquipmentRoom -eq "N") {
                    Send-UserEmail -UserName $UserName -password $UserPassword -requester $Requester -manager $Manager
                }
            }
            if ($O365 -eq "H") {
                Write-Log "Exchange mailbox for $UserName will be created in Exchange Online"
                Write-Log "Calling Create-Mailbox-Hybrid function with the following parameters:"
                Write-Log "UserName: $UserName, realname: $RealName, SharedEquipmentRoom: $SharedEquipmentRoom, Capacity: $Capacity"
                $enabledMailboxes += Create-Mailbox-Hybrid -UserName $UserName -realname $RealName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
                # Send email to requester and manager with new user's name & password
                if ($SharedEquipmentRoom -eq "N") {
                    Send-UserEmail -UserName $UserName -password $UserPassword -requester $Requester -manager $Manager
                }
            }
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
            $MBX = Get-Mailbox $mailbox.alias -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 30
            $i++
        } While (!($MBX) -and $i -lt 5)
        if ($MBX) {
            if ($Mailbox.SharedEquipmentRoom) {
                $logmsg = "Updating Mailbox: " + $Mailbox.Alias +" "+ $Mailbox.SharedEquipmentRoom +" "+ $Mailbox.Capacity +" "
                Write-Log $logMsg
                Update-Mailbox-OnPrem $mailbox.Alias $Mailbox.SharedEquipmentRoom $Mailbox.Capacity
            } elseif (!$Mailbox.SharedEquipmentRoom -and !$Mailbox.Capacity) {
                $logmsg = "Updating Mailbox: " + $Mailbox.Alias
                Write-Log $logMsg
                Update-Mailbox-OnPrem $mailbox.Alias $Mailbox.SharedEquipmentRoom $Mailbox.Capacity
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
    $Err = $null
    Force-ADSync $cred
    $Err = $null
    $connected = $false
    $Failures = @()
    if ($connected -eq $false) {
        try {
            if (!(Get-Module -ListAvailable -Name MSOnline)) {
                Write-Log "Installing MSOnline module"
                Install-Module -Name MSOnline
            }
            if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
                Write-Log "Installing ExchangeOnlineManagement module"
                Install-Module -Name ExchangeOnlineManagement
            }
        } catch {
            Write-Log "EXOv3 PS Module Failed to Install"
            $e = $_.Exception
            Write-Log $e
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log $line
            $msg = $e.Message
            Write-Log $msg
            $Action = "EXOv3 PS Module Failed to Install"
            Write-Log $Action
        }
        try {
            Write-Log "Connecting to Exchange Online"
            Import-Module -Name ExchangeOnlineManagement
            Connect-ExchangeOnline
            Write-Log "Connected to Exchange Online"
            Write-Log "Connecting to Office 365"
            Import-Module -Name MSOnline
            Connect-MsolService
            Write-Log "Connected to Office 365"
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
    if ($connected -eq $false) {
        try {
            Write-Log "Connecting to Exchange Online"
            Import-Module -Name ExchangeOnlineManagement
            Connect-ExchangeOnline
            Write-Log "Connected to Exchange Online"
            Write-Log "Connecting to Office 365"
            Import-Module -Name MSOnline
            Connect-MsolService
            Write-Log "Connected to Office 365"
            $Connected = $true
        } catch {
            $e = $_.Exception
            Write-Log $e
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log $line
            $msg = $e.Message
            Write-Log $msg
            $Action = "Failed to connect via EXOPSSession"
            Write-Log $Action
        }
    }
    if ($connected -eq $true) {
        Write-Log "Updating Mailboxes"
        $LastTry = $True
        Foreach ($mailbox in $EnabledMailboxes) {
            $i = 0
            $MBX = $null
            Do {
                $MBX = Get-Mailbox $mailbox.alias -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 30
                $i++
                if ($i -eq 4 -and $LastTry -eq $True) {
                    Force-ADSync $Cred
                    $LastTry = $False
                    $i = 0
                }
            } While (!($MBX) -and $i -lt 5)
            if ($MBX) {
                if ($Mailbox.SharedEquipmentRoom) {
                    $logmsg = "Updating Mailbox: " + $Mailbox.Alias +" "+ $Mailbox.SharedEquipmentRoom +" "+ $Mailbox.Capacity +" "+ $Mailbox.AdminID
                    Write-Log $logMsg
                    Update-Mailbox-Hybrid $mailbox.Alias $Mailbox.SharedEquipmentRoom $Mailbox.Capacity
                } elseif (!$Mailbox.SharedEquipmentRoom -and !$Mailbox.Capacity) {
                    $logmsg = "Updating Mailbox: " + $Mailbox.Alias
                    Write-Log $logMsg
                    Update-Mailbox-Hybrid $mailbox.Alias $Mailbox.SharedEquipmentRoom $Mailbox.Capacity
                }
            } else {
                $logmsg = "Mailbox: " + $Mailbox.Alias +" not found in AzureAD"
                $Failures += $Mailbox
                Write-Log $logMsg
            }
        }
        foreach ($USER in $CreatedUsers) {
            $UserName = $USER.USERNAME
            $UserName = $UserName.Trim() -replace '[^A-Za-z0-9.]', [String]::Empty # Strip out illegal characters from User ID
            $Dept = $USER.DEPT
            $HiPriv = $USER.HIPRIV.ToUpper()
            $PrivLevel = $USER.PrivLevel
            $SharedEquipmentRoom = $USER.'S/E/R'.ToUpper()
            $UserPrincipalName = "$UserName@$EmailSuffix"
            Write-Log ""
            Write-Log "Assigning region for $UserName"
            Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation GB
            switch ($SharedEquipmentRoom) {
                default {
                    if ($Dept -eq "IT" -and $HiPriv -eq "Y") {
                        $UserNameAdmin = $UserName + ".admin"
                        if ($UserNameAdmin.Length -gt 20) {
                            $UserNameAdmin = $UserNameAdmin.Substring(0,20)
                        }
                        if ($PrivLevel -ge "1") {
                            foreach ($roleName in $Level1Roles) {
                                Write-Log "Assigning roles for $UserNameAdmin"
                                Add-MsolRoleMember -RoleMemberEmailAddress "$UserNameAdmin@$EmailSuffix" -RoleName $roleName
                            }
                        }
                        if ($PrivLevel -ge "2") {
                            foreach ($roleName in $Level2Roles) {
                                Write-Log "Assigning roles for $UserNameAdmin"
                                Add-MsolRoleMember -RoleMemberEmailAddress "$UserNameAdmin@$EmailSuffix" -RoleName $roleName
                            }
                        }
                        if ($PrivLevel -ge "3") {
                            foreach ($roleName in $Level3Roles) {
                                Write-Log "Assigning roles for $UserNameAdmin"
                                Add-MsolRoleMember -RoleMemberEmailAddress "$UserNameAdmin@$EmailSuffix" -RoleName $roleName
                            }
                            $UserNameDomainAdmin = "da." + $UserName
                            if ($UserNameDomainAdmin.Length -gt 20) {
                                $UserNameDomainAdmin = $UserNameDomainAdmin.Substring(0,20)
                            }
                            foreach ($roleName in $Roles) {
                                Write-Log "Assigning roles for $UserNameDomainAdmin"
                                Add-MsolRoleMember -RoleMemberEmailAddress "$UserNameDomainAdmin@$EmailSuffix" -RoleName $roleName
                            }
                        }
                    }
                }
            }
        }
        Disconnect-ExchangeOnline -Confirm:$false
    }
    if (Get-PSSession) {
        Write-Log "Cleaning up PSSessions"
        Get-PSSession | Remove-PSSession
    }
    if ($Failures) {
        if ($Failures.Count -gt 0) {
            Foreach ($Failure in $Failures) {
                $LIST[$LIST.USERNAME.Tolower().IndexOf($Failure.alias.ToLower())] | Export-Csv $FailureFile -NoClobber -NoTypeInformation -Append
            }
        }
    }
}
Write-Log ("=" * 80)
Write-Log "Processing complete"
Write-Log ("=" * 80)
#====================================================================
