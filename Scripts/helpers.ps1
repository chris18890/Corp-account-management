Set-StrictMode -Version Latest

#====================================================================
# Set up logging
#====================================================================
function Write-Log {
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$LogString
        ,[String]$ForegroundColor
    )
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
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$DCHostName
        ,[Parameter(Mandatory)][String]$Group
        ,[Parameter(Mandatory)][String]$Member
        ,[Int]$TimeSpan
    )
    #================================================================
    # Purpose:          To add a user account or group to a group
    # Assumptions:      Parameters have been set correctly
    # Effects:          Member will be added to the group
    # Inputs:           $LogFile - String of log location passed to Write-Log
    #                   $Group - Group name as set before calling the function
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
            Write-Log -LogFile $LogFile -LogString "'$Member' does not exist" -ForegroundColor Red
            return
        }
        Write-Log -LogFile $LogFile -LogString "Adding $Member to $Group"
        try {
            if ($TimeSpan) {
                Add-ADGroupMember -Identity $Group -Members $Member -MemberTimeToLive (New-TimeSpan -Minutes $TimeSpan) -Server $DCHostName
            } else {
                Add-ADGroupMember -Identity $Group -Members $Member -Server $DCHostName
            }
            Write-Log -LogFile $LogFile -LogString "Added $Member to $Group"
        } catch {
            $ex = $_.Exception
            if ($ex.Message -match "already a member") {
                Write-Log -LogFile $LogFile -LogString "'$Member' is already a member of group '$Group'" -ForegroundColor Green
            } else {
                throw
            }
        }
    } else {
        Write-Log -LogFile $LogFile -LogString "$Group does not exist" -ForegroundColor Red
    }
}
#====================================================================

#====================================================================
# Group creation function
#====================================================================
function New-DomainGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$DCHostName
        ,[Parameter(Mandatory)][String]$GroupName
        ,[Parameter(Mandatory)][String]$GroupCategory
        ,[Parameter(Mandatory)][String]$GroupScope
        ,[Parameter(Mandatory)][ValidateSet("E","H","N")][String]$O365
        ,[Parameter(Mandatory)][Boolean]$HiddenFromAddressListsEnabled
        ,[Parameter(Mandatory)][String]$Path
        ,[String]$GroupDescription
    )
    Write-Log -LogFile $LogFile -LogString "Creating Group $GroupName"
    try {
        New-ADGroup -GroupCategory $GroupCategory -GroupScope $GroupScope -Name $GroupName -Path $Path -SamAccountName $GroupName -Server $DCHostName -Description $GroupDescription
        Set-ADObject -Identity "CN=$GroupName,$Path" -Server $DCHostName -ProtectedFromAccidentalDeletion $true
        Write-Log -LogFile $LogFile -LogString "Created $GroupName"
    } catch {
        $ex = $_.Exception
        if ($ex.Message -match "already exists") {
            Write-Log -LogFile $LogFile -LogString "'$GroupName' already exists" -ForegroundColor Green
        } else {
            throw
        }
    }
    if ($O365 -eq "E" -or $O365 -eq "H") {
        try {
            Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
        } catch {
            Write-Log -LogFile $LogFile -LogString "WARNING: Could not enable $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
        }
        try {
            Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $HiddenFromAddressListsEnabled -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
        } catch {
            Write-Log -LogFile $LogFile -LogString "WARNING: Could not configure $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}
#====================================================================

#====================================================================
# Create mailbox function
#====================================================================
function New-UserMailbox {
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$DCHostName
        ,[Parameter(Mandatory)][String]$UserName
        ,[Parameter(Mandatory)][String]$EmailSuffix
        ,[Parameter(Mandatory)][String]$O365EmailSuffix
        ,[String]$realname,[String]$SharedEquipmentRoom,[Int]$Capacity
    )
    #================================================================
    # Purpose:          To create an Exchange Online Mailbox for a user account
    # Assumptions:      Parameters have been set correctly
    # Effects:          Mailbox should be created for user
    # Inputs:           $LogFile - String of log location passed to Write-Log
    #                   $UserName - SAM account name of user
    #                   $realname - Real Name to set as Primary SMTP address, read from CSV
    #                   $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #                   $Capacity - If mailbox is a room account, set the capacity
    # Calls:            Write-Log function
    # Returns:          $EnabledMailbox
    # Notes:            Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-Log -LogFile $LogFile -LogString "Creating mailbox"
    $alias = $UserName.ToUpper()    #Alias = uppercase UserName
    try {
        switch ($SharedEquipmentRoom) {
            "S" {
                if ($realname) {
                    $action = "Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -shared"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -shared
                    }
                } else {
                    $action = "Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -shared"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -shared
                    }
                }
            }
            "E" {
                if ($realname) {
                    $action = "Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -equipment"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -equipment
                    }
                } else {
                    $action = "Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -equipment"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -equipment
                    }
                }
            }
            "R" {
                if ($realname) {
                    $action = "Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -room"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -room
                    }
                } else {
                    $action = "Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix -room"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix" -room
                    }
                }
            }
            default {
                if ($realname) {
                    $action = "Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix"
                    }
                } else {
                    $action = "Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-RemoteMailbox -Identity $UserName -alias $alias -DomainController $DCHostName -remoteroutingaddress "$UserName@$O365EmailSuffix"
                    }
                }
            }
        }
        $mailboxExists = Get-Mailbox $UserName -DomainController $DCHostName -ErrorAction SilentlyContinue
        if (-not $mailboxExists) {
            Write-Log -LogFile $LogFile -LogString "ERROR: Mailbox for $UserName could not be created" -ForegroundColor Red
            throw
        } else {
            Write-Log -LogFile $LogFile -LogString "Mailbox created for $UserName successfully"
            $EnabledMailbox = New-Object -Property @{"Alias" = ""; "SharedEquipmentRoom" = ""; "Capacity" = ""} -TypeName PSObject
            $EnabledMailbox.alias = $alias
            $EnabledMailbox.SharedEquipmentRoom = $SharedEquipmentRoom
            $EnabledMailbox.Capacity = $Capacity
            Return $EnabledMailbox
        }
    } catch {
        $e = $_.Exception
        Write-Log -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-Log -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Failed to enable Mailbox or update settings"
        Write-Log -LogFile $LogFile -LogString $Action -ForegroundColor Red
    }
    Write-Log -LogFile $LogFile -LogString "End of Mailbox Creation Function"
}
#====================================================================

#====================================================================
# Update mailbox Default Settings
#====================================================================
function Update-UserMailbox {
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$UserPrincipalName
        ,[String]$SharedEquipmentRoom = "",[Int]$Capacity = 0
    )
    #================================================================
    # Purpose:          Update Mailbox parameters which need to be configured in O365
    # Assumptions:      Parameters have been set correctly
    # Effects:          Mailbox defaults should be assigned to the new mailbox
    # Inputs:           $LogFile - String of log location passed to Write-Log
    #                   $UserPrincipalName - UPN of user
    #                   $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #                   $Capacity - If mailbox is a room account, set the capacity
    # Calls:            Write-Log function
    # Returns:
    # Notes:            Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-Log -LogFile $LogFile -LogString "Updating Mailbox"
    $MBX = $null
    try {
        $MBX = Get-Mailbox -Identity $UserPrincipalName
        $i = 0
        while (!($MBX) -and ($i -le 6)) {
            $MBX = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
            $i++
            Start-Sleep -seconds 10
        }
        if ($MBX) {
            Write-Log -LogFile $LogFile -LogString " "
            Write-Log -LogFile $LogFile -LogString "Assigning region for $UserPrincipalName"
            Update-MgUser -UserId $UserPrincipalName -UsageLocation GB
            Set-MailboxSpellingConfiguration -Identity $UserPrincipalName -DictionaryLanguage EnglishUnitedKingdom
            Set-MailboxRegionalConfiguration -Identity $UserPrincipalName -Language en-GB -DateFormat "dd/MM/yyyy" -TimeFormat "HH:mm" -TimeZone "GMT Standard Time"
            $identityStr = $UserPrincipalName + ":\Calendar"
            Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Reviewer
            switch ($SharedEquipmentRoom) {
                "S" {
                    if ($MBX.RecipientTypeDetails -ne "SharedMailbox") {
                        Write-Log -LogFile $LogFile -LogString "Converting $UserPrincipalName Mailbox to shared type"
                        Set-Mailbox -Identity $UserPrincipalName -type:shared
                    }
                    Write-Log -LogFile $LogFile -LogString "Updating Shared Mailbox $UserPrincipalName : Adding Permissions"
                    $GroupName = "sh_$UserPrincipalName"
                    Add-MailboxPermission -Identity $UserPrincipalName -User $GroupName -AccessRights FullAccess -confirm:$false
                    Add-RecipientPermission -Identity $UserPrincipalName -Trustee $GroupName -AccessRights "SendAs" -confirm:$false
                    Write-Log -LogFile $LogFile -LogString "Delegated permissions for mailbox $UserPrincipalName to group $GroupName"
                }
                "E" {
                    if ($MBX.RecipientTypeDetails -ne "EquipmentMailbox") {
                        Write-Log -LogFile $LogFile -LogString "Converting $UserPrincipalName Mailbox to equipment type"
                        Set-Mailbox -Identity $UserPrincipalName -type:equipment
                    }
                    #Set Default calendar permissions to Author
                    $MBXType = (Get-Mailbox -Identity $UserPrincipalName).RecipientTypeDetails
                    $x = 0
                    While ($MBXType -ne "EquipmentMailbox" -and $x -lt 6) {
                        Start-Sleep -Seconds 10
                        $MBXType = (Get-Mailbox -Identity $UserPrincipalName).RecipientTypeDetails
                        $x++
                    }
                    $identityStr = $UserPrincipalName + ":\Calendar"
                    Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -confirm:$false
                    #Set calendar resource attendant to auto-accept
                    Write-Log -LogFile $LogFile -LogString "Updating Equipment Mailbox $UserPrincipalName : Updating Calendar Processing"
                    Set-CalendarProcessing -Identity $UserPrincipalName -AutomateProcessing AutoAccept -confirm:$false
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserPrincipalName -ResourceCapacity $Capacity
                    }
                    Write-Log -LogFile $LogFile -LogString "Updating Equipment Mailbox $UserPrincipalName : Adding Permissions"
                    $GroupName = "eq_$UserPrincipalName"
                    Add-MailboxPermission -Identity $UserPrincipalName -User $GroupName -AccessRights FullAccess -confirm:$false
                    Add-RecipientPermission -Identity $UserPrincipalName -Trustee $GroupName -Accessrights "SendAs" -confirm:$false
                    Write-Log -LogFile $LogFile -LogString "Delegated permissions for mailbox $UserPrincipalName to group $GroupName"
                }
                "R" {
                    if ($MBX.RecipientTypeDetails -ne "RoomMailbox") {
                        Write-Log -LogFile $LogFile -LogString "Converting $UserPrincipalName Mailbox to room type"
                        Set-Mailbox -Identity $UserPrincipalName -type:room
                    }
                    #Set Default calendar permissions to Author
                    $MBXType = (Get-Mailbox -Identity $UserPrincipalName).RecipientTypeDetails
                    $x = 0
                    While ($MBXType -ne "RoomMailbox" -and $x -lt 6) {
                        Start-Sleep -Seconds 10
                        $MBXType = (Get-Mailbox -Identity $UserPrincipalName).RecipientTypeDetails
                        $x++
                    }
                    $identityStr = $UserPrincipalName + ":\Calendar"
                    Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -confirm:$false
                    #Set calendar resource attendant to auto-accept
                    Write-Log -LogFile $LogFile -LogString "Updating Room Mailbox $UserPrincipalName : Updating Calendar Processing"
                    Set-CalendarProcessing -Identity $UserPrincipalName -AutomateProcessing AutoAccept -confirm:$false
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserPrincipalName -ResourceCapacity $Capacity
                    }
                    Write-Log -LogFile $LogFile -LogString "Updating Room Mailbox $UserPrincipalName : Adding Permissions"
                    $GroupName = "ro_$UserPrincipalName"
                    Add-MailboxPermission -Identity $UserPrincipalName -User $GroupName -AccessRights FullAccess -confirm:$false
                    Add-RecipientPermission -Identity $UserPrincipalName -Trustee $GroupName -Accessrights "SendAs" -confirm:$false
                    Write-Log -LogFile $LogFile -LogString "Delegated permissions for mailbox $UserPrincipalName to group $GroupName"
                }
            }
        } else {
            $logmsg = "Mailbox: " + $UserPrincipalName +" not found in AzureAD"
            Write-Log -LogFile $LogFile -LogString $LogMsg
        }
    } catch {
        $e = $_.Exception
        Write-Log -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-Log -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Failed to Complete Mailbox Update"
        Write-Log -LogFile $LogFile -LogString $Action -ForegroundColor Red
    }
    Write-Log -LogFile $LogFile -LogString "End of Mailbox Update Function"
}
#====================================================================

#====================================================================
# Create mailbox function
#====================================================================
function New-UserOnPremMailbox {
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$DCHostName
        ,[Parameter(Mandatory)][String]$UserName
        ,[Parameter(Mandatory)][String]$EmailSuffix
        ,[String]$realname,[String]$SharedEquipmentRoom,[Int]$Capacity
    )
    #================================================================
    # Purpose:          To create an Exchange On-Prem Mailbox for a user account
    # Assumptions:      Parameters have been set correctly
    # Effects:          Mailbox should be created for user
    # Inputs:           $LogFile - String of log location passed to Write-Log
    #                   $UserName - SAM account name of user
    #                   $realname - Real Name to set as Primary SMTP address, read from CSV
    #                   $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #                   $Capacity - If mailbox is a room account, set the capacity
    # Calls:            Write-Log function
    # Returns:          $EnabledMailbox
    # Notes:            Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-Log -LogFile $LogFile -LogString "Creating mailbox"
    $alias = $UserName.ToUpper()    #Alias = uppercase UserName
    try {
        switch ($SharedEquipmentRoom) {
            "S" {
                if ($realname) {
                    $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -shared"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -shared
                    }
                } else {
                    $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -shared"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -shared
                    }
                }
            }
            "E" {
                if ($realname) {
                    $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -equipment"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -equipment
                    }
                } else {
                    $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -equipment"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -equipment
                    }
                }
            }
            "R" {
                if ($realname) {
                    $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -room"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -room
                    }
                } else {
                    $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -room"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -room
                    }
                }
            }
            default {
                if ($realname) {
                    $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName
                    }
                } else {
                    $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName"
                    Write-Log -LogFile $LogFile -LogString $Action
                    if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName
                    }
                }
            }
        }
        $mailboxExists = Get-Mailbox $UserName -DomainController $DCHostName -ErrorAction SilentlyContinue
        if (-not $mailboxExists) {
            Write-Log -LogFile $LogFile -LogString "ERROR: Mailbox for $UserName could not be created" -ForegroundColor Red
            throw
        } else {
            Write-Log -LogFile $LogFile -LogString "Mailbox created for $UserName successfully"
            $EnabledMailbox = New-Object -Property @{"Alias" = ""; "SharedEquipmentRoom" = ""; "Capacity" = ""} -TypeName PSObject
            $EnabledMailbox.alias = $alias
            $EnabledMailbox.SharedEquipmentRoom = $SharedEquipmentRoom
            $EnabledMailbox.Capacity = $Capacity
            Return $EnabledMailbox
        }
    } catch {
        $e = $_.Exception
        Write-Log -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-Log -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Failed to enable Mailbox or update settings"
        Write-Log -LogFile $LogFile -LogString $Action -ForegroundColor Red
    }
    Write-Log -LogFile $LogFile -LogString "End of Mailbox Creation Function"
}
#====================================================================

#====================================================================
# Update mailbox Default Settings
#====================================================================
function Update-UserOnPremMailbox {
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$DCHostName
        ,[Parameter(Mandatory)][String]$UserName
        ,[String]$SharedEquipmentRoom = "",[Int]$Capacity = 0
    )
    #================================================================
    # Purpose:          Update Mailbox parameters which need to be configured On-Prem
    # Assumptions:      Parameters have been set correctly
    # Effects:          Mailbox defaults should be assigned to the new mailbox
    # Inputs:           $LogFile - String of log location passed to Write-Log
    #                   $UserName - SAM account name of user
    #                   $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #                   $Capacity - If mailbox is a room account, set the capacity
    # Calls:            Write-Log function
    # Returns:
    # Notes:            Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-Log -LogFile $LogFile -LogString "Updating Mailbox"
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
                        Write-Log -LogFile $LogFile -LogString "Converting $UserName Mailbox to shared type"
                        Set-Mailbox -Identity $UserName -type:shared -DomainController $DCHostName
                    }
                    Write-Log -LogFile $LogFile -LogString "Updating Shared Mailbox $UserName : Adding Permissions"
                    $GroupName = "sh_$UserName"
                    try {
                        Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
                        Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
                    } catch {
                        Write-Log -LogFile $LogFile -LogString "WARNING: Could not enable $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false -DomainController $DCHostName
                    Add-ADPermission -Identity $UserName -User $GroupName -AccessRights ExtendedRight -ExtendedRights "Send As" -confirm:$false -DomainController $DCHostName
                    Write-Log -LogFile $LogFile -LogString "Delegated permissions for mailbox $UserName to group $GroupName"
                }
                "E" {
                    if ($MBX.RecipientTypeDetails -ne "EquipmentMailbox") {
                        Write-Log -LogFile $LogFile -LogString "Converting $UserName Mailbox to equipment type"
                        Set-Mailbox -Identity $UserName -type:equipment -DomainController $DCHostName
                    }
                    Write-Log -LogFile $LogFile -LogString "Updating Equipment Mailbox $UserName : Adding Permissions"
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
                    Write-Log -LogFile $LogFile -LogString "Updating Equipment Mailbox $UserName : Updating Calendar Processing"
                    Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false -DomainController $DCHostName
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserName -ResourceCapacity $Capacity -DomainController $DCHostName
                    }
                    $GroupName = "eq_$UserName"
                    try {
                        Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
                        Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
                    } catch {
                        Write-Log -LogFile $LogFile -LogString "WARNING: Could not enable $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false -DomainController $DCHostName
                    Add-ADPermission -Identity $UserName -User $GroupName -AccessRights ExtendedRight -ExtendedRights "Send As" -confirm:$false -DomainController $DCHostName
                    Write-Log -LogFile $LogFile -LogString "Delegated permissions for mailbox $UserName to group $GroupName"
                }
                "R" {
                    if ($MBX.RecipientTypeDetails -ne "RoomMailbox") {
                        Write-Log -LogFile $LogFile -LogString "Converting $UserName Mailbox to room type"
                        Set-Mailbox -Identity $UserName -type:room -DomainController $DCHostName
                    }
                    Write-Log -LogFile $LogFile -LogString "Updating Room Mailbox $UserName : Adding Permissions"
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
                    Write-Log -LogFile $LogFile -LogString "Updating Room Mailbox $UserName : Updating Calendar Processing"
                    Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false -DomainController $DCHostName
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserName -ResourceCapacity $Capacity -DomainController $DCHostName
                    }
                    $GroupName = "ro_$UserName"
                    try {
                        Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
                        Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
                    } catch {
                        Write-Log -LogFile $LogFile -LogString "WARNING: Could not enable $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false -DomainController $DCHostName
                    Add-ADPermission -Identity $UserName -User $GroupName -AccessRights ExtendedRight -ExtendedRights "Send As" -confirm:$false -DomainController $DCHostName
                    Write-Log -LogFile $LogFile -LogString "Delegated permissions for mailbox $UserName to group $GroupName"
                }
            }
        }
    } catch {
        $e = $_.Exception
        Write-Log -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-Log -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Failed to Complete Mailbox Update"
        Write-Log -LogFile $LogFile -LogString $Action -ForegroundColor Red
    }
    Write-Log -LogFile $LogFile -LogString "End of Mailbox Update Function"
}
#====================================================================

#====================================================================
# Test password against password policy
#====================================================================
function Test-Password {
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$Password
        ,[Parameter(Mandatory)][Int]$PasswordLength
    )
    #================================================================
    # Purpose:          Test password against password policy
    # Assumptions:      Password has been generated with enough characters for required groups
    # Effects:          Password should be valid
    # Inputs:           $LogFile - String of log location passed to Write-Log
    #                   $Password
    #                   $PasswordLength
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
        Write-Log -LogFile $LogFile -LogString "Password validated"
        Write-Log -LogFile $LogFile -LogString " "
    } else {
        Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
        Write-Log -LogFile $LogFile -LogString "ERROR: Password does not comply with the password policy, skipping user" -ForegroundColor Red
        Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
        throw
    }
}
#====================================================================
