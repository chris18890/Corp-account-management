[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
)

#====================================================================
#Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
# ADConnect & Exchange settings
$DCHostName = (Get-ADDomainController).HostName # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ExServer = "$Domain-EXCH.$DNSSuffix" #Remote Exchange PS session
# Get containing folder for script to locate supporting files
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
# Set variables
$ScriptTitle = "$Domain User Mailbox Creation Script"
$EnabledMailboxes = @() # Array to Store Completed Mailbox requests for later enumeration
# File locations
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Log "Creating log folder"
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
#create mailbox function
#====================================================================
function Create-Mailbox-OnPrem {
    param(
    [string]$UserName,[string]$realname
    ,[string]$SharedEquipmentRoom,[string]$Capacity
    )
    #================================================================
    # Purpose:          To create an Exchange 2016 Mailbox for a user account
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
                        Write-log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -shared
                    } else {
                        $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -shared"
                        Write-log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -shared
                    }
                }
                "E" {
                    if ($realname) {
                        $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -equipment"
                        Write-log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -equipment
                    } else {
                        $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -equipment"
                        Write-log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -equipment
                    }
                }
                "R" {
                    if ($realname) {
                        $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName -room"
                        Write-log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName -room
                    } else {
                        $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -room"
                        Write-log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName -room
                    }
                }
                default {
                    if ($realname) {
                        $action = "Enable-Mailbox -Identity $UserName -PrimarySmtpAddress $realname@$EmailSuffix -alias $alias -DomainController $DCHostName"
                        Write-log $action
                        $NewMailbox = Enable-Mailbox -Identity $UserName -PrimarySmtpAddress "$realname@$EmailSuffix" -alias $alias -DomainController $DCHostName
                    } else {
                        $action = "Enable-Mailbox -Identity $UserName -alias $alias -DomainController $DCHostName"
                        Write-log $action
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

Write-Log ("=" * 80)
Write-Log "Log file is '$LogFile'"
Write-Log "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log "DC being used is '$DCHostName'"
Write-Log "Script path is '$ScriptPath'"
Write-Log "$ScriptTitle"
Write-Log "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log ("=" * 80)
Write-Log ""
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
#====================================================================

#====================================================================
#Loop through CSV & create users
#====================================================================
$LIST = @(IMPORT-CSV users.csv)
$CreatedUsers = @()
foreach ($USER in $LIST) {
    $UserName = $USER.USERNAME
    $UserName = $UserName.Trim() -replace '[^A-Za-z0-9.]', [String]::Empty # Strip out illegal characters from User ID
    $Dept = $USER.DEPT
    $HiPriv = $USER.HIPRIV.ToUpper()
    $PrivLevel = $USER.PrivLevel
    $SharedEquipmentRoom = $USER.'S/E/R'.ToUpper()
    $Capacity = $USER.Cap
    $RealName = $USER.REALNAME
    $ExistingUser = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$UserName))"
    if ($ExistingUser) {
        try {
            Write-Log ("=" * 80)
            Write-Log "Processing input file record for $UserName..."
            Write-Log ("=" * 80)
            Write-Log "Exchange mailbox for $UserName will be created in Exchange OnPrem"
            Write-Log "Calling Create-Mailbox-OnPrem function with the following parameters:"
            Write-Log "UserName: $UserName, realname: $RealName, SharedEquipmentRoom: $SharedEquipmentRoom, Capacity: $Capacity"
            $enabledMailboxes += Create-Mailbox-OnPrem -UserName $UserName -realname $RealName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
            switch ($SharedEquipmentRoom) {
                "S" {
                    $GroupName = "sh_$UserName"
                    Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
                    Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
                }
                "E" {
                    $GroupName = "eq_$UserName"
                    Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
                    Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
                }
                "R" {
                    $GroupName = "ro_$UserName"
                    Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
                    Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
                }
                default {
                    if ($Dept -eq "IT" -and $HiPriv -eq "Y") {
                        $UserNameAdmin = $UserName + ".admin"
                        if ($UserNameAdmin.Length -gt 20) {
                            $UserNameAdmin = $UserNameAdmin.Substring(0,20)
                        }
                        $enabledMailboxes += Create-Mailbox-OnPrem -UserName $UserNameAdmin -realname $RealName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
                        if ($PrivLevel -eq "3") {
                            $UserNameDomainAdmin = "da." + $UserName
                            if ($UserNameDomainAdmin.Length -gt 20) {
                                $UserNameDomainAdmin = $UserNameDomainAdmin.Substring(0,20)
                            }
                            $enabledMailboxes += Create-Mailbox-OnPrem -UserName $UserNameDomainAdmin -realname $RealName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
                        }
                    }
                }
            }
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
        if ($Mailbox.SharedEquipmentRoom) {
            $logmsg = "Updating Mailbox:" + $Mailbox.Alias +" "+ $Mailbox.SharedEquipmentRoom +" "+ $Mailbox.Capacity +" "
            Write-log $logMsg
            Update-Mailbox-OnPrem $mailbox.Alias $Mailbox.SharedEquipmentRoom $Mailbox.Capacity
        } elseif (!$Mailbox.SharedEquipmentRoom -and !$Mailbox.Capacity) {
            $logmsg = "Updating Mailbox:" + $Mailbox.Alias
            Write-log $logMsg
            Update-Mailbox-OnPrem $mailbox.Alias $Mailbox.SharedEquipmentRoom $Mailbox.Capacity
        }
    } else {
        $logmsg = "Mailbox:" + $Mailbox.Alias +" not found in AD"
        Write-log $logMsg
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
