#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
)

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
    , [Parameter(Mandatory=$false)] [int]$Capacity = ""
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
            #Set shared account settings
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
# Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
# ADConnect & Exchange settings
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ExServer = "$Domain-EXCH.$DNSSuffix" #Remote Exchange PS session
# Get containing folder for script to locate supporting files
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
# Set variables
$ScriptTitle = "$Domain User Mailbox Creation Script"
$EnabledMailboxes = @() # Array to Store Completed Mailbox requests for later enumeration
# File locations
$LogPath = "$ScriptPath\LogFiles"
$UserInputFile = "$ScriptPath\users.csv"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
$LogFileName = $Domain + "_mailbox_log-$(Get-Date -Format 'yyyyMMdd')"
$LogIndex = 0
while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
    $LogIndex ++
}
$LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
#====================================================================

#====================================================================
# Start of script
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
$requiredGroups = @('ADM_Task_Standard_Account_Admins', 'ADM_Task_SER_Account_Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

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
    $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://$ExServer/PowerShell" -Name ExSession -Credential $cred -ErrorAction Stop -authentication Kerberos
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
#====================================================================

#====================================================================
#Loop through CSV & create users
#====================================================================
# Read input file
Write-Log ("=" * 80)
Write-Log "Reading user data from input file '$UserInputFile'"
Write-Log ("=" * 80)
Write-Log ""
# Read list of users from CSV file ignoring first line
$LIST = @(Import-CSV $UserInputFile)
$RequiredHeaders = @(
    "USERNAME","S/E/R","CAP","REALNAME"
)
$Headers = ($LIST | Select-Object -First 1).PSObject.Properties.Name
foreach ($h in $RequiredHeaders) {
    if ($Headers -notcontains $h) {
        throw "CSV missing required column '$h'"
    }
}
# Process each input file record
foreach ($USER in $LIST) {
    $UserName = $USER.USERNAME
    $UserName = $UserName.Trim() -replace '[^A-Za-z0-9.-]', [String]::Empty # Strip out illegal characters from User ID
    if ($UserName.Length -gt 20) {
        $UserName = $UserName.Substring(0,20)
    }
    $SharedEquipmentRoom = $USER.'S/E/R'.ToUpper()
    [int]$Capacity = $USER.Cap
    $RealName = $USER.REALNAME
    $ExistingUser = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$UserName))" -Server $DCHostName
    if ($ExistingUser) {
        try {
            Write-Log ("=" * 80)
            Write-Log "Processing input file record for $UserName..."
            Write-Log ("=" * 80)
            Write-Log "Exchange mailbox for $UserName will be created in Exchange OnPrem"
            Write-Log "Calling New-UserOnPremMailbox function with the following parameters:"
            Write-Log "UserName: $UserName, realname: $RealName, SharedEquipmentRoom: $SharedEquipmentRoom, Capacity: $Capacity"
            $EnabledMailboxes += New-UserOnPremMailbox -UserName $UserName -realname $RealName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
            switch ($SharedEquipmentRoom) {
                "S" {
                    $GroupName = "sh_$UserName"
                    try {
                        Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
                        Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
                    } catch {
                        Write-Log "WARNING: Could not enable $GroupName — $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }
                "E" {
                    $GroupName = "eq_$UserName"
                    try {
                        Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
                        Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
                    } catch {
                        Write-Log "WARNING: Could not enable $GroupName — $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }
                "R" {
                    $GroupName = "ro_$UserName"
                    try {
                        Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
                        Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
                    } catch {
                        Write-Log "WARNING: Could not enable $GroupName — $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }
            }
            Write-Log ("=" * 80)
            Write-Log "Processing input file record for $UserName complete"
            Write-Log ("=" * 80)
            Write-Log ""
        } catch {
            $e = $_.Exception
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log ("-" * 80) -ForegroundColor Red
            Write-Log "ERROR: Failed during processing of $UserName - Line $Line" -ForegroundColor Red
            Write-Log "$e"
            Write-Log ("-" * 80) -ForegroundColor Red
            Write-Log ("=" * 80)
            Write-Log ""
        }
    }
}
Write-Log "Updating Mailboxes"
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
            $logmsg = "Updating Mailbox:" + $Mailbox.Alias +" "+ $Mailbox.SharedEquipmentRoom +" "+ $Mailbox.Capacity +" "
            Write-Log $logMsg
            Update-UserOnPremMailbox $mailbox.Alias $Mailbox.SharedEquipmentRoom $Mailbox.Capacity
        } elseif (!$Mailbox.SharedEquipmentRoom -and !$Mailbox.Capacity) {
            $logmsg = "Updating Mailbox:" + $Mailbox.Alias
            Write-Log $logMsg
            Update-UserOnPremMailbox $mailbox.Alias $Mailbox.SharedEquipmentRoom $Mailbox.Capacity
        }
    } else {
        $logmsg = "Mailbox:" + $Mailbox.Alias +" not found in AD"
        Write-Log $logMsg
    }
}
if (Get-PSSession -Name ExSession -ErrorAction SilentlyContinue) {
    Remove-PsSession $ExSession
    Write-Log "Closed Exchange session."
}
Write-Log ("=" * 80)
Write-Log "Processing complete"
Write-Log ("=" * 80)
#====================================================================
