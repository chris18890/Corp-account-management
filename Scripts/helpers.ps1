Set-StrictMode -Version Latest

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
        , [int]$TimeSpan
    )
    #================================================================
    # Purpose:          To add a user account or group to a group
    # Assumptions:      Parameters have been set correctly
    # Effects:          Member will be added to the group
    # Inputs:           $Group - Group name as set before calling the function
    #                   $Member - Object to be added
    #                   $TimeSpan - number of minutes to add temporal memebership for
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
            if ($TimeSpan) {
                Write-Log "Adding $Member to $Group for $TimeSpan minutes" -ForegroundColor Yellow
                Add-ADGroupMember -Identity $Group -Members $Member -MemberTimeToLive (New-TimeSpan -Minutes $TimeSpan) -Server $DCHostName
            } else {
                Write-Log "Adding $Member to $Group with no time limit, manual removal required" -ForegroundColor Yellow
                Add-ADGroupMember -Identity $Group -Members $Member -Server $DCHostName
            }
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
# Group creation function
#====================================================================
function New-DomainGroup {
    [CmdletBinding()]
    param(
        [string]$GroupName,[String]$GroupScope,[ValidateSet("E","H","N")][string]$O365,[boolean]$HiddenFromAddressListsEnabled,[String]$Path,[String]$GroupDescription
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
# Create mailbox function
#====================================================================
function New-UserMailbox {
    param(
    [string]$UserName
    ,[string]$realname,[string]$SharedEquipmentRoom,[int]$Capacity
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
    Write-Log "Creating mailbox"
    $alias = $UserName.ToUpper()    #Alias = uppercase UserName
    try {
        switch ($SharedEquipmentRoom) {
            "S" {
                if ($realname) {
                    $action = "Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -shared"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -shared
                    }
                } else {
                    $action = "Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -shared"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -shared
                    }
                }
            }
            "E" {
                if ($realname) {
                    $action = "Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -equipment"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -equipment
                    }
                } else {
                    $action = "Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -equipment"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -equipment
                    }
                }
            }
            "R" {
                if ($realname) {
                    $action = "Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -room"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -room
                    }
                } else {
                    $action = "Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -room"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -room
                    }
                }
            }
            default {
                if ($realname) {
                    $action = "Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix"
                    }
                } else {
                    $action = "Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix"
                    }
                }
            }
        }
        $mailboxExists = Get-Mailbox $UserName -DomainController $DCHostName -ErrorAction SilentlyContinue
        if (-not $mailboxExists) {
            Write-Log "ERROR: Mailbox for $UserName could not be created" -ForegroundColor Red
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
#====================================================================

#====================================================================
# Update mailbox Default Settings
#====================================================================
function Update-UserMailbox {
    param(
    [Parameter(Mandatory=$true)] [string]$UserName
    , [Parameter(Mandatory=$false)] [string]$SharedEquipmentRoom = ""
    , [Parameter(Mandatory=$false)] [int]$Capacity = 0
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
    $alias = $UserName.ToUpper()    #Alias = uppercase UserName
    $MBX = $null
    try {
        $MBX = Get-Mailbox -Identity $UserName
        $i = 0
        while (!($MBX) -and ($i -le 6)) {
            $MBX = Get-Mailbox -Identity $UserName -ErrorAction SilentlyContinue
            $i++
            Start-Sleep -seconds 10
        }
        if ($MBX) {
            Set-MailboxSpellingConfiguration -Identity $UserName -DictionaryLanguage EnglishUnitedKingdom
            Set-MailboxRegionalConfiguration -Identity $UserName -Language en-GB -DateFormat "dd/MM/yyyy" -TimeFormat "HH:mm" -TimeZone "GMT Standard Time"
            $identityStr = $UserName + ":\Calendar"
            Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Reviewer
            switch ($SharedEquipmentRoom) {
                "S" {
                    if ($MBX.RecipientTypeDetails -ne "SharedMailbox") {
                        Write-Log "Converting $UserName Mailbox to shared type"
                        Set-Mailbox -Identity $UserName -type:shared
                    }
                    Write-Log "Updating Shared Mailbox $UserName : Adding Permissions"
                    $GroupName = "sh_$UserName@$EmailSuffix"
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false
                    Add-RecipientPermission -Identity $UserName -trustee $GroupName -Accessrights "SendAs" -confirm:$false
                    Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                }
                "E" {
                    if ($MBX.RecipientTypeDetails -ne "EquipmentMailbox") {
                        Write-Log "Converting $UserName Mailbox to equipment type"
                        Set-Mailbox -Identity $UserName -type:equipment
                    }
                    #Set Default calendar permissions to Author
                    $MBXType = (Get-Mailbox -Identity $UserName).RecipientTypeDetails
                    $x = 0
                    While ($MBXType -ne "EquipmentMailbox" -and $x -lt 6) {
                        Start-Sleep -Seconds 10
                        $MBXType = (Get-Mailbox -Identity $UserName).RecipientTypeDetails
                        $x++
                    }
                    $identityStr = $UserName + ":\Calendar"
                    Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -confirm:$false
                    #Set calendar resource attendant to auto-accept
                    Write-Log "Updating Equipment Mailbox $UserName : Updating Calendar Processing"
                    Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserName -ResourceCapacity $Capacity
                    }
                    Write-Log "Updating Equipment Mailbox $UserName : Adding Permissions"
                    $GroupName = "eq_$UserName@$EmailSuffix"
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false
                    Add-RecipientPermission -Identity $UserName -Trustee $GroupName -Accessrights "SendAs" -confirm:$false
                    Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                }
                "R" {
                    if ($MBX.RecipientTypeDetails -ne "RoomMailbox") {
                        Write-Log "Converting $UserName Mailbox to room type"
                        Set-Mailbox -Identity $UserName -type:room
                    }
                    #Set Default calendar permissions to Author
                    $MBXType = (Get-Mailbox -Identity $UserName).RecipientTypeDetails
                    $x = 0
                    While ($MBXType -ne "RoomMailbox" -and $x -lt 6) {
                        Start-Sleep -Seconds 10
                        $MBXType = (Get-Mailbox -Identity $UserName).RecipientTypeDetails
                        $x++
                    }
                    $identityStr = $UserName + ":\Calendar"
                    Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -confirm:$false
                    #Set calendar resource attendant to auto-accept
                    Write-Log "Updating Room Mailbox $UserName : Updating Calendar Processing"
                    Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserName -ResourceCapacity $Capacity
                    }
                    Write-Log "Updating Room Mailbox $UserName : Adding Permissions"
                    $GroupName = "ro_$UserName@$EmailSuffix"
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false
                    Add-RecipientPermission -Identity $UserName -Trustee $GroupName -Accessrights "SendAs" -confirm:$false
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
    Write-Log "End of Mailbox Update Function"
}
#====================================================================

#====================================================================
# Create mailbox function
#====================================================================
function New-UserOnPremMailbox {
    param(
    [string]$UserName
    ,[string]$realname,[string]$SharedEquipmentRoom,[int]$Capacity
    )
    #================================================================
    # Purpose:          To create an Exchange On-Prem Mailbox for a user account
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
    Write-Log "Creating mailbox"
    $alias = $UserName.ToUpper()    #Alias = uppercase UserName
    try {
        switch ($SharedEquipmentRoom) {
            "S" {
                if ($realname) {
                    $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -shared"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -shared
                    }
                } else {
                    $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -shared"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -shared
                    }
                }
            }
            "E" {
                if ($realname) {
                    $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -equipment"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -equipment
                    }
                } else {
                    $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -equipment"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -equipment
                    }
                }
            }
            "R" {
                if ($realname) {
                    $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -room"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -room
                    }
                } else {
                    $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -room"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -room
                    }
                }
            }
            default {
                if ($realname) {
                    $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName
                    }
                } else {
                    $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName"
                    Write-Log $action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName
                    }
                }
            }
        }
        $mailboxExists = Get-Mailbox $UserName -DomainController $DCHostName -ErrorAction SilentlyContinue
        if (-not $mailboxExists) {
            Write-Log "ERROR: Mailbox for $UserName could not be created" -ForegroundColor Red
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
#====================================================================

#====================================================================
# Update mailbox Default Settings
#====================================================================
function Update-UserOnPremMailbox {
    param(
    [Parameter(Mandatory=$true)] [string]$UserName
    , [Parameter(Mandatory=$false)] [string]$SharedEquipmentRoom = ""
    , [Parameter(Mandatory=$false)] [int]$Capacity = 0
    )
    #================================================================
    # Purpose:          Update Mailbox parameters which need to be configured On-Prem
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
    $alias = $UserName.ToUpper()    #Alias = uppercase UserName
    $MBX = $null
    try {
        $MBX = Get-Mailbox -Identity $UserName -DomainController $DCHostName
        $i = 0
        while (!($MBX) -and ($i -le 6)) {
            $MBX = Get-Mailbox -Identity $UserName -DomainController $DCHostName -erroraction silentlycontinue
            $i++
            Start-Sleep -seconds 10
        }
        if ($MBX) {
            Set-MailboxSpellingConfiguration -Identity $UserName -DictionaryLanguage EnglishUnitedKingdom -DomainController $DCHostName
            Set-MailboxRegionalConfiguration -Identity $UserName -Language en-GB -DateFormat "dd/MM/yyyy" -TimeFormat "HH:mm" -TimeZone "GMT Standard Time" -DomainController $DCHostName
            $identityStr = $UserName + ":\Calendar"
            Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Reviewer -DomainController $DCHostName
            switch ($SharedEquipmentRoom) {
                "S" {
                    if ($MBX.RecipientTypeDetails -ne "SharedMailbox") {
                        Write-Log "Converting $UserName Mailbox to shared type"
                        Set-Mailbox -Identity $UserName -type:shared -DomainController $DCHostName
                    }
                    Write-Log "Updating Shared Mailbox $UserName : Adding Permissions"
                    $GroupName = "sh_$UserName"
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false -DomainController $DCHostName
                    Add-ADPermission -Identity $UserName -User $GroupName -ExtendedRights "Send As" -confirm:$false -DomainController $DCHostName
                    Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                }
                "E" {
                    if ($MBX.RecipientTypeDetails -ne "EquipmentMailbox") {
                        Write-Log "Converting $UserName Mailbox to equipment type"
                        Set-Mailbox -Identity $UserName -type:equipment -DomainController $DCHostName
                    }
                    Write-Log "Updating Equipment Mailbox $UserName : Adding Permissions"
                    #Set Default calendar permissions to Author
                    $MBXType = (Get-Mailbox -Identity $UserName -DomainController $DCHostName).RecipientTypeDetails
                    $x = 0
                    While ($MBXType -ne "EquipmentMailbox" -and $x -lt 6) {
                        Start-Sleep -Seconds 10
                        $MBXType = (Get-Mailbox -Identity $UserName -DomainController $DCHostName).RecipientTypeDetails
                        $x++
                    }
                    $identityStr = $UserName + ":\Calendar"
                    Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -confirm:$false -DomainController $DCHostName
                    #Set calendar resource attendant to auto-accept
                    Write-Log "Updating Equipment Mailbox $UserName : Updating Calendar Processing"
                    Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false -DomainController $DCHostName
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserName -ResourceCapacity $Capacity -DomainController $DCHostName
                    }
                    $GroupName = "eq_$UserName"
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false -DomainController $DCHostName
                    Add-ADPermission -Identity $UserName -User $GroupName -ExtendedRights "Send As" -confirm:$false -DomainController $DCHostName
                    Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                }
                "R" {
                    if ($MBX.RecipientTypeDetails -ne "RoomMailbox") {
                        Write-Log "Converting $UserName Mailbox to room type"
                        Set-Mailbox -Identity $UserName -type:room -DomainController $DCHostName
                    }
                    Write-Log "Updating Room Mailbox $UserName : Adding Permissions"
                    #Set Default calendar permissions to Author
                    $MBXType = (Get-Mailbox -Identity $UserName -DomainController $DCHostName).RecipientTypeDetails
                    $x = 0
                    While ($MBXType -ne "RoomMailbox" -and $x -lt 6) {
                        Start-Sleep -Seconds 10
                        $MBXType = (Get-Mailbox -Identity $UserName -DomainController $DCHostName).RecipientTypeDetails
                        $x++
                    }
                    $identityStr = $UserName + ":\Calendar"
                    Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -confirm:$false -DomainController $DCHostName
                    #Set calendar resource attendant to auto-accept
                    Write-Log "Updating Room Mailbox $UserName : Updating Calendar Processing"
                    Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false -DomainController $DCHostName
                    if ($Capacity) {
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
    Write-Log "End of Mailbox Update Function"
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
    # Notes:            There are 5 requirements in the current policy, but this could change in future
    #================================================================
    $TestsPassed = 0
    if ($Password.length -ge ($PasswordLength)) {$TestsPassed ++} # Must be >= 20 characters in length
    if ($Password -cmatch "[a-z]") {$TestsPassed ++} # Must contain a lowercase letter
    if ($Password -cmatch "[A-Z]") {$TestsPassed ++} # Must contain an uppercase letter
    if ($Password -cmatch "[0-9]") {$TestsPassed ++} # Must contain a digit
    if ($Password -cmatch "[^a-zA-Z0-9]") {$TestsPassed ++} # Must contain a special character
    if ($TestsPassed -ge 5) {
        Write-Log "Password validated"
        Write-Log ""
    } else {
        Write-Log ("-" * 80) -ForegroundColor Red
        Write-Log "ERROR: Password does not comply with the password policy, skipping user" -ForegroundColor Red
        Write-Log ("-" * 80) -ForegroundColor Red
        throw
    }
}
#====================================================================
