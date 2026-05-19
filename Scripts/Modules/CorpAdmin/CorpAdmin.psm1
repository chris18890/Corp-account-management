Set-StrictMode -Version Latest

# Capture the directory containing corpadmin.psm1 at dot-source time so the
# default environment.psd1 path resolves regardless of the caller's location.
$Script:HelpersDir = $PSScriptRoot
$Script:CachedEnv = $null
$Script:CachedPath = $null

#====================================================================
# Load the centralised environment configuration
#====================================================================
function Get-EnvironmentConfig {
    [CmdletBinding()]
    param(
        [string]$Path
        ,[switch]$Force
    )
    # Resolve the effective path FIRST, so the cache can be keyed on it.
    if (-not $Path) {
        # Allow ops override: set $env:CORPADMIN_ENV_PSD1 to point at an
        # alternate config file (useful for multi-tenant deployments).
        if ($env:CORPADMIN_ENV_PSD1 -and (Test-Path -LiteralPath $env:CORPADMIN_ENV_PSD1)) {
            $Path = $env:CORPADMIN_ENV_PSD1
        }
        else {
            # Module lives at Scripts/Modules/CorpAdmin/; environment.psd1 lives at Scripts/.
            $Path = Join-Path (Split-Path (Split-Path $Script:HelpersDir -Parent) -Parent) 'environment.psd1'
        }
    }
    # Cache hit only when not forced AND the resolved path matches what we cached.
    # This means an env-var change (or unset) between calls correctly invalidates.
    if (-not $Force -and $Script:CachedEnv -and $Script:CachedPath -eq $Path) {
        return $Script:CachedEnv
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "environment.psd1 not found at '$Path'. Pass -Path explicitly."
    }
    $config = Import-PowerShellDataFile -LiteralPath $Path
    # Structural validation: fail fast and loudly if a required top-level
    # section is absent, rather than letting a downstream script dereference
    # a null section (e.g. $Env.WSUS.Products) and fail far from the cause.
    $requiredSections = @(
        'Network','OUs','Groups','Shares','Locale',
        'Security','Azure','Exchange','EntraRoles','WSUS'
    )
    $missing = $requiredSections | Where-Object {
        -not $config.ContainsKey($_) -or $null -eq $config[$_]
    }
    if ($missing) {
        throw "environment.psd1 at '$Path' is missing required section(s): $($missing -join ', ')."
    }
    # Cache every resolution, keyed on the path it came from. An explicit
    # -Path call now also seeds the cache (keyed on that path) rather than
    # deliberately not populating it - which is fine, because the key
    # guarantees a later default-path call with a different resolved path
    # won't get this entry.
    $Script:CachedEnv  = $config
    $Script:CachedPath = $Path
    $config
}
#====================================================================

#====================================================================
# Set up logging
#====================================================================
function Write-LogFile {
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
    $line = "$(Get-Date -Format 'G') $LogString"
    $line | Out-File -FilePath $LogFile -Append -Encoding UTF8
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
    # Inputs:           $LogFile - String of log location passed to Write-LogFile
    #                   $Group - Group name as set before calling the function
    #                   $Member - Object to be added
    #                   $TimeSpan - number of minutes to add temporal memebership for
    # Calls:            Write-LogFile function
    # Returns:
    # Notes:
    #================================================================
    $checkGroup = Get-ADGroup -LDAPFilter "(SAMAccountName=$Group)" -Server $DCHostName
    if ($null -ne $checkGroup) {
        $checkMember = Get-ADObject -LDAPFilter "(SAMAccountName=$Member)" -Server $DCHostName
        if (-not $checkMember) {
            Write-LogFile -LogFile $LogFile -LogString "'$Member' does not exist" -ForegroundColor Red
            throw "'$Member' does not exist"
        }
        try {
            Write-LogFile -LogFile $LogFile -LogString "Adding $Member to $Group"
            if ($TimeSpan) {
                Add-ADGroupMember -Identity $Group -Members $Member -MemberTimeToLive (New-TimeSpan -Minutes $TimeSpan) -Server $DCHostName
            } else {
                Add-ADGroupMember -Identity $Group -Members $Member -Server $DCHostName
            }
            Write-LogFile -LogFile $LogFile -LogString "Added $Member to $Group"
        } catch {
            $ex = $_.Exception
            if ($ex.Message -match "already a member") {
                Write-LogFile -LogFile $LogFile -LogString "'$Member' is already a member of group '$Group'" -ForegroundColor Green
            } else {
                throw
            }
        }
    } else {
        Write-LogFile -LogFile $LogFile -LogString "$Group does not exist" -ForegroundColor Red
        throw "Group '$Group' does not exist"
    }
}
#====================================================================

#====================================================================
# OU creation function
#====================================================================
function New-ADOU {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$DCHostName
        ,[Parameter(Mandatory)][string]$OUName
        ,[Parameter(Mandatory)][String]$Path
        ,[String]$OUDescription
    )
    Write-LogFile -LogFile $LogFile -LogString "Creating OU $OUName"
    try {
        New-ADOrganizationalUnit -Name $OUName -Path $Path -ProtectedFromAccidentalDeletion:$true -Description $OUDescription -Server $DCHostName
        Write-LogFile -LogFile $LogFile -LogString "Created OU $OUName"
    } catch {
        $ex = $_.Exception
        if ($ex.Message -match "already in use") {
            Write-LogFile -LogFile $LogFile -LogString "'$OUName' already exists" -ForegroundColor Green
        } else {
            throw
        }
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
    Write-LogFile -LogFile $LogFile -LogString "Creating Group $GroupName"
    try {
        New-ADGroup -GroupCategory $GroupCategory -GroupScope $GroupScope -Name $GroupName -Path $Path -SamAccountName $GroupName -Server $DCHostName -Description $GroupDescription
        Set-ADObject -Identity "CN=$GroupName,$Path" -Server $DCHostName -ProtectedFromAccidentalDeletion $true
        Write-LogFile -LogFile $LogFile -LogString "Created Group $GroupName"
    } catch {
        $ex = $_.Exception
        if ($ex.Message -match "already exists") {
            Write-LogFile -LogFile $LogFile -LogString "'$GroupName' already exists" -ForegroundColor Green
        } else {
            throw
        }
    }
    if ($O365 -eq "E" -or $O365 -eq "H") {
        try {
            Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "WARNING: Could not enable $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
        }
        try {
            Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $HiddenFromAddressListsEnabled -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "WARNING: Could not configure $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
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
    # Inputs:           $LogFile - String of log location passed to Write-LogFile
    #                   $UserName - SAM account name of user
    #                   $realname - Real Name to set as Primary SMTP address, read from CSV
    #                   $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #                   $Capacity - If mailbox is a room account, set the capacity
    # Calls:            Write-LogFile function
    # Returns:          $EnabledMailbox
    # Notes:            Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-LogFile -LogFile $LogFile -LogString "Creating mailbox"
    $alias = $UserName.ToUpper()    #Alias = uppercase UserName
    try {
        $mbxParams = @{
            Identity             = $UserName
            alias                = $alias
            DomainController     = $DCHostName
            remoteroutingaddress = "$UserName@$O365EmailSuffix"
        }
        if ($realname) {
            $mbxParams['PrimarySmtpAddress'] = "$realname@$EmailSuffix"
        }
        switch ($SharedEquipmentRoom) {
            "S" { $mbxParams['shared']    = $true }
            "E" { $mbxParams['equipment'] = $true }
            "R" { $mbxParams['room']      = $true }
        }
        $smtp = if ($realname) { " -PrimarySmtpAddress $realname@$EmailSuffix" } else { "" }
        $flag = switch ($SharedEquipmentRoom) { "S" { " -shared" } "E" { " -equipment" } "R" { " -room" } default { "" } }
        $action = "Enable-RemoteMailbox -Identity $UserName$smtp -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix$flag"
        Write-LogFile -LogFile $LogFile -LogString $Action
        if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
            Enable-RemoteMailbox @mbxParams
        }
        $mailboxExists = Get-Mailbox $UserName -DomainController $DCHostName -ErrorAction SilentlyContinue
        if (-not $mailboxExists) {
            Write-LogFile -LogFile $LogFile -LogString "ERROR: Mailbox for $UserName could not be created" -ForegroundColor Red
            throw
        } else {
            Write-LogFile -LogFile $LogFile -LogString "Mailbox created for $UserName successfully"
            $EnabledMailbox = New-Object -Property @{"Alias" = ""; "SharedEquipmentRoom" = ""; "Capacity" = ""} -TypeName PSObject
            $EnabledMailbox.alias = $alias
            $EnabledMailbox.SharedEquipmentRoom = $SharedEquipmentRoom
            $EnabledMailbox.Capacity = $Capacity
            Return $EnabledMailbox
        }
    } catch {
        $e = $_.Exception
        Write-LogFile -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-LogFile -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-LogFile -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Failed to enable Mailbox or update settings"
        Write-LogFile -LogFile $LogFile -LogString $Action -ForegroundColor Red
    }
    Write-LogFile -LogFile $LogFile -LogString "End of Mailbox Creation Function"
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
    # Inputs:           $LogFile - String of log location passed to Write-LogFile
    #                   $UserPrincipalName - UPN of user
    #                   $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #                   $Capacity - If mailbox is a room account, set the capacity
    # Calls:            Write-LogFile function
    # Returns:
    # Notes:            Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-LogFile -LogFile $LogFile -LogString "Updating Mailbox"
    $MBX = $null
    try {
        $Env = Get-EnvironmentConfig
        $MBX = Get-Mailbox -Identity $UserPrincipalName
        $i = 0
        while (!($MBX) -and ($i -le 6)) {
            $MBX = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
            $i++
            Start-Sleep -seconds 10
        }
        if ($MBX) {
            Write-LogFile -LogFile $LogFile -LogString " "
            Write-LogFile -LogFile $LogFile -LogString "Assigning region for $UserPrincipalName"
            Update-MgUser -UserId $UserPrincipalName -UsageLocation $Env.Locale.UsageLocation
            Set-MailboxSpellingConfiguration -Identity $UserPrincipalName -DictionaryLanguage $Env.Locale.Dictionary
            Set-MailboxRegionalConfiguration -Identity $UserPrincipalName -Language $Env.Locale.Language -DateFormat $Env.Locale.DateFormat -TimeFormat $Env.Locale.TimeFormat -TimeZone $Env.Locale.TimeZone
            $identityStr = $UserPrincipalName + ":\Calendar"
            Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Reviewer
            switch ($SharedEquipmentRoom) {
                "S" {
                    if ($MBX.RecipientTypeDetails -ne "SharedMailbox") {
                        Write-LogFile -LogFile $LogFile -LogString "Converting $UserPrincipalName Mailbox to shared type"
                        Set-Mailbox -Identity $UserPrincipalName -type:shared
                    }
                    Write-LogFile -LogFile $LogFile -LogString "Updating Shared Mailbox $UserPrincipalName : Adding Permissions"
                    $GroupName = "$($Env.Groups.SharedAccessPrefix)$UserPrincipalName"
                    Add-MailboxPermission -Identity $UserPrincipalName -User $GroupName -AccessRights FullAccess -confirm:$false
                    Add-RecipientPermission -Identity $UserPrincipalName -Trustee $GroupName -AccessRights "SendAs" -confirm:$false
                    Write-LogFile -LogFile $LogFile -LogString "Delegated permissions for mailbox $UserPrincipalName to group $GroupName"
                }
                "E" {
                    if ($MBX.RecipientTypeDetails -ne "EquipmentMailbox") {
                        Write-LogFile -LogFile $LogFile -LogString "Converting $UserPrincipalName Mailbox to equipment type"
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
                    Write-LogFile -LogFile $LogFile -LogString "Updating Equipment Mailbox $UserPrincipalName : Updating Calendar Processing"
                    Set-CalendarProcessing -Identity $UserPrincipalName -AutomateProcessing AutoAccept -confirm:$false
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserPrincipalName -ResourceCapacity $Capacity
                    }
                    Write-LogFile -LogFile $LogFile -LogString "Updating Equipment Mailbox $UserPrincipalName : Adding Permissions"
                    $GroupName = "$($Env.Groups.EquipmentAccessPrefix)$UserPrincipalName"
                    Add-MailboxPermission -Identity $UserPrincipalName -User $GroupName -AccessRights FullAccess -confirm:$false
                    Add-RecipientPermission -Identity $UserPrincipalName -Trustee $GroupName -Accessrights "SendAs" -confirm:$false
                    Write-LogFile -LogFile $LogFile -LogString "Delegated permissions for mailbox $UserPrincipalName to group $GroupName"
                }
                "R" {
                    if ($MBX.RecipientTypeDetails -ne "RoomMailbox") {
                        Write-LogFile -LogFile $LogFile -LogString "Converting $UserPrincipalName Mailbox to room type"
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
                    Write-LogFile -LogFile $LogFile -LogString "Updating Room Mailbox $UserPrincipalName : Updating Calendar Processing"
                    Set-CalendarProcessing -Identity $UserPrincipalName -AutomateProcessing AutoAccept -confirm:$false
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserPrincipalName -ResourceCapacity $Capacity
                    }
                    Write-LogFile -LogFile $LogFile -LogString "Updating Room Mailbox $UserPrincipalName : Adding Permissions"
                    $GroupName = "$($Env.Groups.RoomAccessPrefix)$UserPrincipalName"
                    Add-MailboxPermission -Identity $UserPrincipalName -User $GroupName -AccessRights FullAccess -confirm:$false
                    Add-RecipientPermission -Identity $UserPrincipalName -Trustee $GroupName -Accessrights "SendAs" -confirm:$false
                    Write-LogFile -LogFile $LogFile -LogString "Delegated permissions for mailbox $UserPrincipalName to group $GroupName"
                }
            }
        } else {
            $logmsg = "Mailbox: " + $UserPrincipalName +" not found in AzureAD"
            Write-LogFile -LogFile $LogFile -LogString $LogMsg
        }
    } catch {
        $e = $_.Exception
        Write-LogFile -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-LogFile -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-LogFile -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Failed to Complete Mailbox Update"
        Write-LogFile -LogFile $LogFile -LogString $Action -ForegroundColor Red
    }
    Write-LogFile -LogFile $LogFile -LogString "End of Mailbox Update Function"
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
    # Inputs:           $LogFile - String of log location passed to Write-LogFile
    #                   $UserName - SAM account name of user
    #                   $realname - Real Name to set as Primary SMTP address, read from CSV
    #                   $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #                   $Capacity - If mailbox is a room account, set the capacity
    # Calls:            Write-LogFile function
    # Returns:          $EnabledMailbox
    # Notes:            Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-LogFile -LogFile $LogFile -LogString "Creating mailbox"
    $alias = $UserName.ToUpper()    #Alias = uppercase UserName
    try {
        $mbxParams = @{
            Identity         = $UserName
            alias            = $alias
            DomainController = $DCHostName
        }
        if ($realname) {
            $mbxParams['PrimarySmtpAddress'] = "$realname@$EmailSuffix"
        }
        switch ($SharedEquipmentRoom) {
            "S" { $mbxParams['shared']    = $true }
            "E" { $mbxParams['equipment'] = $true }
            "R" { $mbxParams['room']      = $true }
        }
        $smtp = if ($realname) { " -PrimarySmtpAddress $realname@$EmailSuffix" } else { "" }
        $flag = switch ($SharedEquipmentRoom) { "S" { " -shared" } "E" { " -equipment" } "R" { " -room" } default { "" } }
        $action = "Enable-Mailbox -Identity $UserName$smtp -alias $alias -DomainController $DCHostName$flag"
        Write-LogFile -LogFile $LogFile -LogString $Action
        if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
            Enable-Mailbox @mbxParams
        }
        $mailboxExists = Get-Mailbox $UserName -DomainController $DCHostName -ErrorAction SilentlyContinue
        if (-not $mailboxExists) {
            Write-LogFile -LogFile $LogFile -LogString "ERROR: Mailbox for $UserName could not be created" -ForegroundColor Red
            throw
        } else {
            Write-LogFile -LogFile $LogFile -LogString "Mailbox created for $UserName successfully"
            $EnabledMailbox = New-Object -Property @{"Alias" = ""; "SharedEquipmentRoom" = ""; "Capacity" = ""} -TypeName PSObject
            $EnabledMailbox.alias = $alias
            $EnabledMailbox.SharedEquipmentRoom = $SharedEquipmentRoom
            $EnabledMailbox.Capacity = $Capacity
            Return $EnabledMailbox
        }
    } catch {
        $e = $_.Exception
        Write-LogFile -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-LogFile -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-LogFile -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Failed to enable Mailbox or update settings"
        Write-LogFile -LogFile $LogFile -LogString $Action -ForegroundColor Red
    }
    Write-LogFile -LogFile $LogFile -LogString "End of Mailbox Creation Function"
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
    # Inputs:           $LogFile - String of log location passed to Write-LogFile
    #                   $UserName - SAM account name of user
    #                   $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #                   $Capacity - If mailbox is a room account, set the capacity
    # Calls:            Write-LogFile function
    # Returns:
    # Notes:            Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-LogFile -LogFile $LogFile -LogString "Updating Mailbox"
    $MBX = $null
    try {
        $Env = Get-EnvironmentConfig
        $MBX = Get-Mailbox -Identity $UserName -DomainController $DCHostName
        $i = 0
        while (!($MBX) -and ($i -le 6)) {
            $MBX = Get-Mailbox -Identity $UserName -DomainController $DCHostName -erroraction silentlycontinue
            $i++
            Start-Sleep -seconds 10
        }
        if ($MBX) {
            Set-MailboxSpellingConfiguration -Identity $UserName -DictionaryLanguage $Env.Locale.Dictionary -DomainController $DCHostName
            Set-MailboxRegionalConfiguration -Identity $UserName -Language $Env.Locale.Language -DateFormat $Env.Locale.DateFormat -TimeFormat $Env.Locale.TimeFormat -TimeZone $Env.Locale.TimeZone -DomainController $DCHostName
            $identityStr = $UserName + ":\Calendar"
            Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Reviewer -DomainController $DCHostName
            switch ($SharedEquipmentRoom) {
                "S" {
                    if ($MBX.RecipientTypeDetails -ne "SharedMailbox") {
                        Write-LogFile -LogFile $LogFile -LogString "Converting $UserName Mailbox to shared type"
                        Set-Mailbox -Identity $UserName -type:shared -DomainController $DCHostName
                    }
                    Write-LogFile -LogFile $LogFile -LogString "Updating Shared Mailbox $UserName : Adding Permissions"
                    $GroupName = "$($Env.Groups.SharedAccessPrefix)$UserName"
                    try {
                        Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
                        Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "WARNING: Could not enable $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false -DomainController $DCHostName
                    Add-ADPermission -Identity $UserName -User $GroupName -AccessRights ExtendedRight -ExtendedRights "Send As" -confirm:$false -DomainController $DCHostName
                    Write-LogFile -LogFile $LogFile -LogString "Delegated permissions for mailbox $UserName to group $GroupName"
                }
                "E" {
                    if ($MBX.RecipientTypeDetails -ne "EquipmentMailbox") {
                        Write-LogFile -LogFile $LogFile -LogString "Converting $UserName Mailbox to equipment type"
                        Set-Mailbox -Identity $UserName -type:equipment -DomainController $DCHostName
                    }
                    Write-LogFile -LogFile $LogFile -LogString "Updating Equipment Mailbox $UserName : Adding Permissions"
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
                    Write-LogFile -LogFile $LogFile -LogString "Updating Equipment Mailbox $UserName : Updating Calendar Processing"
                    Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false -DomainController $DCHostName
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserName -ResourceCapacity $Capacity -DomainController $DCHostName
                    }
                    $GroupName = "$($Env.Groups.EquipmentAccessPrefix)$UserName"
                    try {
                        Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
                        Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "WARNING: Could not enable $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false -DomainController $DCHostName
                    Add-ADPermission -Identity $UserName -User $GroupName -AccessRights ExtendedRight -ExtendedRights "Send As" -confirm:$false -DomainController $DCHostName
                    Write-LogFile -LogFile $LogFile -LogString "Delegated permissions for mailbox $UserName to group $GroupName"
                }
                "R" {
                    if ($MBX.RecipientTypeDetails -ne "RoomMailbox") {
                        Write-LogFile -LogFile $LogFile -LogString "Converting $UserName Mailbox to room type"
                        Set-Mailbox -Identity $UserName -type:room -DomainController $DCHostName
                    }
                    Write-LogFile -LogFile $LogFile -LogString "Updating Room Mailbox $UserName : Adding Permissions"
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
                    Write-LogFile -LogFile $LogFile -LogString "Updating Room Mailbox $UserName : Updating Calendar Processing"
                    Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false -DomainController $DCHostName
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserName -ResourceCapacity $Capacity -DomainController $DCHostName
                    }
                    $GroupName = "$($Env.Groups.RoomAccessPrefix)$UserName"
                    try {
                        Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
                        Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "WARNING: Could not enable $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false -DomainController $DCHostName
                    Add-ADPermission -Identity $UserName -User $GroupName -AccessRights ExtendedRight -ExtendedRights "Send As" -confirm:$false -DomainController $DCHostName
                    Write-LogFile -LogFile $LogFile -LogString "Delegated permissions for mailbox $UserName to group $GroupName"
                }
            }
        }
    } catch {
        $e = $_.Exception
        Write-LogFile -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-LogFile -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-LogFile -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Failed to Complete Mailbox Update"
        Write-LogFile -LogFile $LogFile -LogString $Action -ForegroundColor Red
    }
    Write-LogFile -LogFile $LogFile -LogString "End of Mailbox Update Function"
}
#====================================================================

#====================================================================
# Bias-free random index in [0, $max) via rejection sampling.
# uint64 throughout so the threshold doesn't overflow when $max divides 2^32.
#====================================================================
function Get-CryptoRandIndex {
    param([int]$max)
    $rng = $null
    if ($max -le 1) { return 0 }
    try {
        $rng   = [System.Security.Cryptography.RandomNumberGenerator]::Create()
        $bytes = [byte[]]::new(4)
        $rangeSize = [uint64]4294967296   # 2^32
        $threshold = $rangeSize - ($rangeSize % [uint64]$max)
        do {
            $rng.GetBytes($bytes)
            $r = [uint64][BitConverter]::ToUInt32($bytes, 0)
        } while ($r -ge $threshold)
        [int]($r % [uint64]$max)
    } finally {
        if ($rng) {
            $rng.Dispose()
        }
    }
}
#====================================================================

#====================================================================
# Generate a cryptographically random password 
#====================================================================
function New-Password {
    <#
    .SYNOPSIS
        Generate a cryptographically random password.
    .DESCRIPTION
        Drop-in replacement for [Web.Security.Membership]::GeneratePassword.
        Uses System.Security.Cryptography.RandomNumberGenerator with rejection
        sampling so each character draw is unbiased.
        
        Character sets exclude visually ambiguous characters (I, O, l, 0, 1)
        to reduce transcription errors when passwords are handed off to users
        verbally or in writing. Special character set avoids quote, backslash,
        pipe, backtick, slash and whitespace.
        
        Guarantees:
        - At least $MinLower lowercase, $MinUpper uppercase, $MinDigit digit,
          and $MinSpecial special characters
        - Fisher-Yates shuffle so required-class positions aren't predictable
        
        Compatible with Windows PowerShell 5.1 and PowerShell 7+; does not
        rely on [RandomNumberGenerator]::GetInt32 (.NET Core / 5+ only).
    .EXAMPLE
        $pw = New-Password
    .EXAMPLE
        $pw = New-Password -Length 32 -MinSpecial 6
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [ValidateRange(12, 256)][int]$Length     = 20
        ,[ValidateRange(0, 64)][int]$MinUpper    = 1
        ,[ValidateRange(0, 64)][int]$MinLower    = 1
        ,[ValidateRange(0, 64)][int]$MinDigit    = 1
        ,[ValidateRange(0, 64)][int]$MinSpecial  = 4
    )
    $minTotal = $MinUpper + $MinLower + $MinDigit + $MinSpecial
    if ($minTotal -gt $Length) {
        throw "Minimum character requirements ($minTotal) exceed total password length ($Length)."
    }
    # Ambiguous chars removed: I O l 0 1
    $upper   = [char[]]'ABCDEFGHJKLMNPQRSTUVWXYZ'
    $lower   = [char[]]'abcdefghijkmnopqrstuvwxyz'
    $digit   = [char[]]'23456789'
    $special = [char[]]'!@#$%^&*()-_=+[]{};:,.<>?'
    $all     = $upper + $lower + $digit + $special
    $chars = [System.Collections.Generic.List[char]]::new()
    # Required minimums
    for ($i = 0; $i -lt $MinUpper;   $i++) { $chars.Add($upper[(Get-CryptoRandIndex $upper.Length)]) }
    for ($i = 0; $i -lt $MinLower;   $i++) { $chars.Add($lower[(Get-CryptoRandIndex $lower.Length)]) }
    for ($i = 0; $i -lt $MinDigit;   $i++) { $chars.Add($digit[(Get-CryptoRandIndex $digit.Length)]) }
    for ($i = 0; $i -lt $MinSpecial; $i++) { $chars.Add($special[(Get-CryptoRandIndex $special.Length)]) }
    # Fill the rest from the union
    while ($chars.Count -lt $Length) {
        $chars.Add($all[(Get-CryptoRandIndex $all.Length)])
    }
    # Fisher-Yates shuffle so the required-class chars aren't at fixed positions
    for ($i = $chars.Count - 1; $i -gt 0; $i--) {
        $j = Get-CryptoRandIndex ($i + 1)
        $tmp = $chars[$i]; $chars[$i] = $chars[$j]; $chars[$j] = $tmp
    }
    -join $chars
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
    # Inputs:           $LogFile - String of log location passed to Write-LogFile
    #                   $Password
    #                   $PasswordLength
    # Calls:            Write-LogFile function
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
        Write-LogFile -LogFile $LogFile -LogString "Password validated"
        Write-LogFile -LogFile $LogFile -LogString " "
    } else {
        Write-LogFile -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
        Write-LogFile -LogFile $LogFile -LogString "ERROR: Password does not comply with the password policy, skipping user" -ForegroundColor Red
        Write-LogFile -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
        throw "Password does not comply with the password policy"
    }
}
#====================================================================

#====================================================================
# GPO link function
#====================================================================
function Add-GPOLink {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$DCHostName
        ,[Parameter(Mandatory)][string]$GPOName
        ,[Parameter(Mandatory)][string]$GPOTarget
    )
    #================================================================
    # Purpose:          To link a GPO to an OU
    # Assumptions:      Parameters have been set correctly
    # Effects:          GPO will be linked to the OU
    # Inputs:           $GPOName - Name of GPO as set before calling the function
    #                   $GPOTarget - OU where GPO will be linked
    # Returns:
    # Notes:
    #================================================================
    Write-LogFile -LogFile $LogFile -LogString "Linking $GPOName to $GPOTarget"
    try {
        New-GPLink -name $GPOName -target $GPOTarget -LinkEnabled Yes -enforced yes -Order 1 -ErrorAction Stop -Server $DCHostName
        Write-LogFile -LogFile $LogFile -LogString "Linked $GPOName to $GPOTarget"
    } catch {
        $ex = $_.Exception
        if ($ex.Message -match "already linked") {
            Write-LogFile -LogFile $LogFile -LogString "'$GPOName' already linked to $GPOTarget" -ForegroundColor Green
        } else {
            throw
        }
    }
}
#====================================================================

#====================================================================
# AD Sync
#====================================================================
function Invoke-ADSync {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][pscredential]$Cred
        ,[Parameter(Mandatory)][String]$AzureADConnect
        ,[Parameter(Mandatory)][String]$O365EmailSuffix
    )
    try {
        $ADConnectSession = New-PSSession -Computername $AzureADConnect -Credential $Cred
        Invoke-Command -Session $ADConnectSession {Import-Module ADSync}
        Import-PSSession -Session $ADConnectSession -Module ADSync -AllowClobber
        $state = (Get-ADSyncConnectorRunStatus | Where-Object { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
        $ADSyncLoop = 0
        while ($State -and $ADSyncLoop -le 10) {
            Write-LogFile -LogFile $LogFile -LogString "AD Sync Connector is currently busy, waiting 30 seconds before trying again"
            Start-Sleep -Seconds 30
            $State = (Get-ADSyncConnectorRunStatus | Where-Object { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
            $ADSyncLoop++
        }
        if ($ADSyncLoop -ge 10) {
            Write-LogFile -LogFile $LogFile -LogString "AD Sync Connector has returned a busy state for 5 minutes or more, if this continues, please contact the servicedesk to investigate further"
        } else {
            Write-LogFile -LogFile $LogFile -LogString "Attempting to run Azure AD Sync Cycle"
            Start-ADSyncSyncCycle -PolicyType Delta
        }
        $state = (Get-ADSyncConnectorRunStatus | Where-Object { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
        $ADSyncLoop = 0
        while ($State -and $ADSyncLoop -le 10) {
            Write-LogFile -LogFile $LogFile -LogString "AD Sync Connector is busy, waiting 30 seconds To allow sync to complete"
            Start-Sleep -Seconds 30
            $State = (Get-ADSyncConnectorRunStatus | Where-Object { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
            $ADSyncLoop++
        }
        if (!($state) -and $ADSyncLoop -le 10) {
            Write-LogFile -LogFile $LogFile -LogString "AD Sync complete"
        } else {
            Write-LogFile -LogFile $LogFile -LogString "AD Sync has not completed within 5 minutes, please check log for issues relating to syncronization issues."
        }
    } catch {
        $e = $_.Exception
        Write-LogFile -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-LogFile -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-LogFile -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Unable to Sync AD"
        Write-LogFile -LogFile $LogFile -LogString $Action -ForegroundColor Red
    } finally {
        if ($ADConnectSession) { Remove-PSSession $ADConnectSession }
    }
}
#====================================================================

#====================================================================
# Build the schema GUID map (lDAPDisplayName -> schemaIDGUID).
# Used by callers needing GUIDs for AccessRule construction.
#====================================================================
function Get-ADSchemaGuidMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server
    )
    $rootdse = Get-ADRootDSE -Server $Server
    $map = @{}
    $params = @{
        SearchBase = $rootdse.SchemaNamingContext
        LDAPFilter = '(schemaidguid=*)'
        Properties = ('lDAPDisplayName', 'schemaIDGUID')
    }
    Get-ADObject @params -Server $Server | ForEach-Object {
        $map[$_.lDAPDisplayName] = [System.GUID]$_.schemaIDGUID
    }
    return $map
}
#====================================================================

#====================================================================
# Build the extended rights map (displayName -> rightsGuid).
#====================================================================
function Get-ADExtendedRightsMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server
    )
    $rootdse = Get-ADRootDSE -Server $Server
    $map = @{}
    $params = @{
        SearchBase = $rootdse.ConfigurationNamingContext
        LDAPFilter = '(&(objectclass=controlAccessRight)(rightsguid=*))'
        Properties = ('displayName', 'rightsGuid')
    }
    Get-ADObject @params -Server $Server | ForEach-Object {
        $map[$_.displayName] = [System.GUID]$_.rightsGuid
    }
    return $map
}
#====================================================================

#====================================================================
# Delegate permission on computer objects to a group.
# Grants: CreateChild, DeleteChild, WriteProperty, WriteDacl, plus
# validated writes (DNS host name, SPN) and Reset/Change Password
# extended rights, all scoped to computer objects under $TargetOU.
#====================================================================
function Grant-ComputerJoinDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetOU
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $Acl = Get-Acl "AD:\OU=$TargetOU,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'CreateChild',$AccessControlTypeAllow,$GuidMap['computer'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'DeleteChild',$AccessControlTypeAllow,$GuidMap['computer'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,'Descendents',$GuidMap['computer']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteDacl',$AccessControlTypeAllow,'Descendents',$GuidMap['computer']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'Self',$AccessControlTypeAllow,$ExtendedRightsMap['Validated write to DNS host name'],'Descendents',$GuidMap['computer']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'Self',$AccessControlTypeAllow,$ExtendedRightsMap['Validated write to service principal name'],'Descendents',$GuidMap['computer']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ExtendedRightAccessRule $AdminGroupSID,$AccessControlTypeAllow,$ExtendedRightsMap['Reset Password'],'Descendents',$GuidMap['computer']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ExtendedRightAccessRule $AdminGroupSID,$AccessControlTypeAllow,$ExtendedRightsMap['Change Password'],'Descendents',$GuidMap['computer']))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permission on group objects to a group.
# Grants: CreateChild, DeleteChild, WriteProperty on group objects.
#====================================================================
function Grant-GroupDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetOU
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $Acl = Get-Acl "AD:\OU=$TargetOU,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'CreateChild',$AccessControlTypeAllow,$GuidMap['group'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'DeleteChild',$AccessControlTypeAllow,$GuidMap['group'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,'Descendents',$GuidMap['group']))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate modify membership permission on group objects to a group.
# Grants: WriteProperty on the 'member' attribute of group objects.
#====================================================================
function Grant-GroupMembershipEditDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetOU
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $Acl = Get-Acl "AD:\OU=$TargetOU,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,$GuidMap['member'],'Descendents',$GuidMap['group']))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate password reset permission on user objects to a group.
# Grants: WriteProperty on pwdLastSet / lockoutTime, plus Reset Password
# extended right, on user objects under $TargetOU.
#====================================================================
function Grant-PasswordResetDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetOU
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $Acl = Get-Acl "AD:\OU=$TargetOU,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,$GuidMap['pwdLastSet'],'Descendents',$GuidMap['user']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,$GuidMap['lockoutTime'],'Descendents',$GuidMap['user']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'ExtendedRight',$AccessControlTypeAllow,$ExtendedRightsMap['Reset Password'],'Descendents',$GuidMap['user']))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permission on user objects to a group.
# Grants: CreateChild, DeleteChild, WriteProperty, plus Reset Password
# extended right, on user objects under $TargetOU.
#====================================================================
function Grant-UserDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetOU
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $Acl = Get-Acl "AD:\OU=$TargetOU,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'CreateChild',$AccessControlTypeAllow,$GuidMap['user'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'DeleteChild',$AccessControlTypeAllow,$GuidMap['user'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,'Descendents',$GuidMap['user']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'ExtendedRight',$AccessControlTypeAllow,$ExtendedRightsMap['Reset Password'],'Descendents',$GuidMap['user']))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permission on OU objects to a group.
# Grants: CreateChild, DeleteChild, WriteProperty on OU objects.
#====================================================================
function Grant-OUDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetOU
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $Acl = Get-Acl "AD:\OU=$TargetOU,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'CreateChild',$AccessControlTypeAllow,$GuidMap['organizationalUnit'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'DeleteChild',$AccessControlTypeAllow,$GuidMap['organizationalUnit'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,'Descendents',$GuidMap['organizationalUnit']))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to DNS Operators.
# Grants: GenericRead, GenericExecute, GenericWrite, CreateChild,
# DeleteChild Allow on the DNS container.
#====================================================================
function Grant-DNSOperatorsPermissionDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetDN
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $adRightsGR = [System.DirectoryServices.ActiveDirectoryRights] 'GenericRead'
    $adRightsGE = [System.DirectoryServices.ActiveDirectoryRights] 'GenericExecute'
    $adRightsGW = [System.DirectoryServices.ActiveDirectoryRights] 'GenericWrite'
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] 'CreateChild'
    $adRightsDC = [System.DirectoryServices.ActiveDirectoryRights] 'DeleteChild'
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] 'All'
    $Acl = Get-Acl "AD:\$TargetDN,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGR,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGE,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGW,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsCC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsDC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to DNS ReadOnly group.
# Grants: GenericRead, GenericExecute Allow.
# Denies: GenericWrite, CreateChild, DeleteChild, WriteOwner, WriteDacl,
# DeleteTree, Delete.
# Then strips the implicit Deny on ReadControl (which GenericWrite bit
# would otherwise drag in) so the group can still inspect the ACL.
#====================================================================
function Grant-DNSReadOnlyPermissionDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetDN
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $adRightsGR = [System.DirectoryServices.ActiveDirectoryRights] 'GenericRead'
    $adRightsGE = [System.DirectoryServices.ActiveDirectoryRights] 'GenericExecute'
    $adRightsGW = [System.DirectoryServices.ActiveDirectoryRights] 'GenericWrite'
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] 'CreateChild'
    $adRightsDC = [System.DirectoryServices.ActiveDirectoryRights] 'DeleteChild'
    $adRightsRC = [System.DirectoryServices.ActiveDirectoryRights] 'ReadControl'
    $adRightsWO = [System.DirectoryServices.ActiveDirectoryRights] 'WriteOwner'
    $adRightsWD = [System.DirectoryServices.ActiveDirectoryRights] 'WriteDacl'
    $adRightsDT = [System.DirectoryServices.ActiveDirectoryRights] 'DeleteTree'
    $adRightsDEL = [System.DirectoryServices.ActiveDirectoryRights] 'Delete'
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $AccessControlTypeDeny = [System.Security.AccessControl.AccessControlType] 'Deny'
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] 'All'
    $Acl = Get-Acl "AD:\$TargetDN,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGR,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGE,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGW,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsCC,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsDC,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWO,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWD,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsDT,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsDEL,$AccessControlTypeDeny,$inheritanceTypeAll))
    # GenericWrite includes the ReadControl bit, so its Deny ACE implicitly denies ReadControl too; strip that so the group can still inspect the ACL.
    $Acl.RemoveAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsRC,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate AD Sites/Subnets/Transports admin permission to a group.
# Grants: GenericAll, CreateChild, DeleteChild Allow on $TargetDN.
#====================================================================
function Grant-ADObjectPermissionDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetDN
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $adRightsGA = [System.DirectoryServices.ActiveDirectoryRights] 'GenericAll'
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] 'CreateChild'
    $adRightsDC = [System.DirectoryServices.ActiveDirectoryRights] 'DeleteChild'
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] 'All'
    $Acl = Get-Acl "AD:\$TargetDN,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGA,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsCC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsDC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate GPO link/option/RSoP permissions to a group.
# Grants: ReadProperty + WriteProperty on gPLink and gPOptions schema
# attributes plus the two RSoP extended rights, on the domain root.
#====================================================================
function Grant-GPOPermissionDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $adRightsRP = [System.DirectoryServices.ActiveDirectoryRights] 'ReadProperty'
    $adRightsWP = [System.DirectoryServices.ActiveDirectoryRights] 'WriteProperty'
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] 'All'
    $Acl = Get-Acl "AD:\$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsRP,$AccessControlTypeAllow,$GuidMap['gPLink'],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWP,$AccessControlTypeAllow,$GuidMap['gPLink'],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsRP,$AccessControlTypeAllow,$GuidMap['gPOptions'],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWP,$AccessControlTypeAllow,$GuidMap['gPOptions'],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'ExtendedRight',$AccessControlTypeAllow,$ExtendedRightsMap['Generate Resultant Set of Policy (Logging)'],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'ExtendedRight',$AccessControlTypeAllow,$ExtendedRightsMap['Generate Resultant Set of Policy (Planning)'],$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to create GPOs.
# Grants: CreateChild on $TargetDN with no inheritance.
#====================================================================
function Grant-GPOCreationDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetDN
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] 'CreateChild'
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $inheritanceTypeNone = [System.DirectoryServices.ActiveDirectorySecurityInheritance] 'None'
    $Acl = Get-Acl "AD:\$TargetDN,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsCC,$AccessControlTypeAllow,$inheritanceTypeNone))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
function ConvertTo-IntOrDefault {
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [AllowNull()][AllowEmptyString()][string]$Value
        ,[int]$Default = 0
    )
    $result = $Default
    if (-not [string]::IsNullOrWhiteSpace($Value)) {
        if (-not [int]::TryParse($Value.Trim(), [ref]$result)) {
            $result = $Default
        }
    }
    $result
}
#====================================================================

#====================================================================
function ConvertTo-SafeName {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)][AllowEmptyString()][string]$Value
    )
    process {
        # Strip characters illegal in Office 365 name fields: ? @ \ +
        # (\\ is an escaped backslash inside the regex character class).
        ("$Value").Trim() -replace '[?@\\+]', [String]::Empty
    }
}
#====================================================================

#====================================================================
function ConvertTo-SafeSamAccountName {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)][AllowEmptyString()][string]$Value
        # Trusted prefix applied AFTER the value is sanitised but BEFORE
        # truncation (mirrors CreateITAdminUser's "admin." + $UserName then cap).
        ,[Parameter()][string]$Prefix = ''
        ,[Parameter()][ValidateRange(1, 256)][int]$MaxLength = 20
    )
    process {
        # Keep only SAM-safe characters: letters, digits, dot, hyphen.
        $clean = $Prefix + (("$Value").Trim() -replace '[^A-Za-z0-9.-]', [String]::Empty)
        if ($clean.Length -gt $MaxLength) {
            $clean = $clean.Substring(0, $MaxLength)
        }
        $clean
    }
}
#====================================================================

#====================================================================
function Send-NotificationEmail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][string]$SMTPServer
        ,[Parameter(Mandatory)][string]$EmailTo
        ,[Parameter(Mandatory)][string]$EmailFrom
        ,[Parameter(Mandatory)][string]$EmailSubject
        ,[Parameter(Mandatory)][string]$EmailBody
    )
    #================================================================
    # Purpose:          To send an email
    # Assumptions:      Parameters have been set correctly
    # Effects:          Email will be sent
    # Inputs:           $LogFile
    #                   $EmailTo
    #                   $EmailFrom
    #                   $EmailBody
    #                   $EmailSubject
    #                   $SMTPServer
    # Calls:            Write-LogFile function
    # Returns:
    # Notes:
    #================================================================
    Import-Module Send-MailKitMessage
    $RecipientList = [MimeKit.InternetAddressList]::new();
    $RecipientList.Add([MimeKit.InternetAddress]$EmailTo);
    $Splat = @{
        RecipientList   = $RecipientList
        From            = $EmailFrom
        Body            = $EmailBody
        Subject         = $EmailSubject
        SmtpServer      = $SMTPServer
        UseSecureConnectionIfAvailable = $true
    }
    try {
       Send-MailKitMessage @Splat
       Write-LogFile -LogFile $LogFile -LogString "Notification email sent to $EmailTo"
    } catch {
        Write-LogFile -LogFile $LogFile -LogString "ERROR sending notification email to $EmailTo : $_" -ForegroundColor Red
    }
}
#====================================================================
