[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
)

$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptTitle = "$Domain User Sync Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Log "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
$LogFileName = $Domain + "_new_user_cleanup_log-$(Get-Date -Format 'yyyyMMdd')"
$LogIndex = 0
while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
    $LogIndex ++
}
$LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
$Roles = @("Company Administrator")
$Level1Roles = @("Helpdesk Administrator", "Service support administrator", "Global Reader")
$Level2Roles = @("User Administrator", "Groups administrator", "Authentication administrator", "License Administrator")
$Level3Roles = @("Exchange Administrator", "Teams Administrator", "Sharepoint Administrator", "Privileged authentication administrator", "Privileged role administrator")

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

Write-Log ("=" * 80)
Write-Log "Log file is '$LogFile'"
Write-Log "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log "Script path is '$ScriptPath'"
Write-Log "$ScriptTitle"
Write-Log "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log ("=" * 80)
Write-Log ""

#====================================================================
#Loop through CSV & create users
#====================================================================
$CreatedUsers = @(IMPORT-CSV "users.csv")
try {
    if (!(Get-Module -ListAvailable -Name MSOnline)) {
        Write-Log "Installing MSOnline module"
        Install-Module -Name MSOnline
    }
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log "Installing ExchangeOnlineManagement module"
        Install-Module -Name ExchangeOnlineManagement
    }
    Write-Log "Starting AzureAD Sync"
    Import-Module -Name ADSync
    Start-ADSyncSyncCycle -PolicyType Delta
    Write-Log "Connecting to Office 365"
    Import-Module -Name MSOnline
    Connect-MsolService
    Write-Log "Connected to Office 365"
    Write-Log "Connecting to Exchange Online"
    Import-Module -Name ExchangeOnlineManagement
    Connect-ExchangeOnline
    Write-Log "Connected to Exchange Online"
    Write-Log "Pausing for 30 seconds"
    Start-Sleep -s 30
    foreach ($USER in $CreatedUsers) {
        $UserName = $USER.USERNAME
        $UserName = $UserName.Trim() -replace '[^A-Za-z0-9.]', [String]::Empty # Strip out illegal characters from User ID
        $Dept = $USER.DEPT
        $HiPriv = $USER.HIPRIV.ToUpper()
        $PrivLevel = $USER.PrivLevel
        $SharedEquipmentRoom = $USER.'S/E/R'.ToUpper()
        $Capacity = $USER.Cap
        $UserPrincipalName = "$UserName@$EmailSuffix"
        $MBX = $null
        $MBX = Get-Mailbox -Identity $UserPrincipalName
        $i = 0
        while (!($MBX) -and ($i -le 6)) {
            $MBX = Get-Mailbox -Identity $UserPrincipalName -erroraction silentlycontinue
            $i++
            Start-Sleep -seconds 10
        }
        if ($MBX) {
            Write-Log ""
            Write-Log "Assigning region for $UserName"
            Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation GB
            Set-MailboxSpellingConfiguration -Identity $UserPrincipalName -DictionaryLanguage EnglishUnitedKingdom
            Set-MailboxRegionalConfiguration -Identity $UserPrincipalName -Language en-GB -DateFormat "dd/MM/yyyy" -TimeFormat "HH:mm" -TimeZone "GMT Standard Time"
            $identityStr = $UserPrincipalName + ":\Calendar"
            Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Reviewer
            switch ($SharedEquipmentRoom) {
                "S" {
                    if ($MBX.RecipientTypeDetails -ne "SharedMailbox") {
                        Write-Log "Converting $UserName Mailbox to shared type"
                        Set-Mailbox -Identity $UserPrincipalName -type:shared
                    }
                    Write-Log "Updating Shared Mailbox $UserName : Adding Permissions"
                    $GroupName = "sh_$UserName"
                    Add-MailboxPermission -Identity $UserPrincipalName -User $GroupName -AccessRights FullAccess -InheritanceType All -confirm:$false
                    Add-RecipientPermission -Identity $UserPrincipalName -Trustee $GroupName -AccessRights SendAs -confirm:$false
                    Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                }
                "E" {
                    if ($MBX.RecipientTypeDetails -ne "EquipmentMailbox") {
                        Write-Log "Converting $UserName Mailbox to equipment type"
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
                    Set-CalendarProcessing -Identity $UserPrincipalName -AutomateProcessing AutoAccept -confirm:$false
                    Write-Log "Updating Equipment Mailbox $UserName : Adding Permissions"
                    $GroupName = "eq_$UserName"
                    Add-MailboxPermission -Identity $UserPrincipalName -User $GroupName -AccessRights FullAccess -confirm:$false
                    Add-RecipientPermission -Identity $UserPrincipalName -Trustee $GroupName -Accessrights "SendAs" -confirm:$false
                    Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                }
                "R" {
                    if ($MBX.RecipientTypeDetails -ne "RoomMailbox") {
                        Write-Log "Converting $UserName Mailbox to room type"
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
                    Set-CalendarProcessing -Identity $UserPrincipalName -AutomateProcessing AutoAccept -confirm:$false
                    if ($Capacity -ne "N") {
                        Set-Mailbox -Identity $UserPrincipalName -ResourceCapacity $Capacity
                    }
                    Write-Log "Updating Room Mailbox $UserName : Adding Permissions"
                    $GroupName = "ro_$UserName"
                    Add-MailboxPermission -Identity $UserPrincipalName -User $GroupName -AccessRights FullAccess -confirm:$false
                    Add-RecipientPermission -Identity $UserPrincipalName -Trustee $GroupName -Accessrights "SendAs" -confirm:$false
                    Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                }
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
        } else {
            $logmsg = "Mailbox: " + $UserName +" not found in AzureAD"
            Write-Log $logMsg
        }
    }
    Write-Log ""
    Write-Log "Office 365 sync & mailbox update complete"
    Disconnect-ExchangeOnline -Confirm:$false
} catch {
    $e = $_.Exception
    $line = $_.InvocationInfo.ScriptLineNumber
    $msg = $e.Message
    $Action = "Failed to Complete Mailbox Update"
    Write-Log $e $line $msg $Action
    Write-Log $Action
}
#====================================================================
