[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
)

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
$Level3Roles = @("Exchange Administrator", "Teams Administrator", "Sharepoint Administrator", "Privileged authentication administrator")

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
$CreatedUsers = @(IMPORT-CSV users.csv)
try {
    if (!(Get-Module -ListAvailable -Name MSOnline)) {
        Write-Log "Installing MSOnline module"
        Install-Module MSOnline
    }
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log "Installing ExchangeOnlineManagement module"
        Install-Module -Name ExchangeOnlineManagement
    }
    Write-Log "Starting AzureAD Sync"
    Import-Module ADSync
    Start-ADSyncSyncCycle -PolicyType Delta
    Write-Log "Connecting to Office 365"
    Connect-MsolService
    Write-Log "Connecting to Exchange Online"
    Connect-ExchangeOnline
    Write-Log "Pausing for 30 seconds"
    Start-Sleep -s 30
    foreach ($USER in $CreatedUsers) {
        $UserName = $USER.USERNAME
        $UserName = $UserName.Trim() -replace '[^A-Za-z0-9]', [String]::Empty # Strip out illegal characters from User ID
        $Dept = $USER.DEPT
        $HiPriv = $USER.HIPRIV.ToUpper()
        $PrivLevel = $USER.PrivLevel
        $SharedEquipmentRoom = $USER.'S/E/R'.ToUpper()
        $Capacity = $USER.Cap
        $UserPrincipalName = "$UserName@$EmailSuffix"
        $MBX = $null
        $MBX = Get-Mailbox -Identity $UserName
        $i = 0
        while (!($MBX) -and ($i -le 6)) {
            $MBX = Get-Mailbox -Identity $UserName -erroraction silentlycontinue
            $i++
            Start-Sleep -seconds 10
        }
        if ($MBX) {
            Write-Log ""
            Write-Log "Assigning region for $UserName"
            Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation GB
            Set-MailboxSpellingConfiguration -Identity $UserName -DictionaryLanguage EnglishUnitedKingdom
            Set-MailboxRegionalConfiguration -Identity $UserName -Language en-GB -DateFormat "dd/MM/yyyy" -TimeFormat "HH:mm" -TimeZone "GMT Standard Time"
            $identityStr = $UserName + ":\Calendar"
            Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Reviewer
            switch ($SharedEquipmentRoom) {
                "S" {
                    Write-Log "Updating Shared Mailbox $UserName : Adding Permissions"
                    $GroupName = "sh_$UserName"
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -InheritanceType All -confirm:$false
                    Add-RecipientPermission -Identity $UserName -Trustee $GroupName -AccessRights SendAs -confirm:$false
                    Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                }
                "E" {
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
                    Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false
                    Write-Log "Updating Equipment Mailbox $UserName : Adding Permissions"
                    $GroupName = "eq_$UserName"
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false
                    Add-RecipientPermission -Identity $UserName -Trustee $GroupName -Accessrights "SendAs" -confirm:$false
                    Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                }
                "R" {
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
                    Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -confirm:$false
                    if ($Capacity -ne "N") {
                        Set-Mailbox -Identity $UserName -ResourceCapacity $Capacity
                    }
                    Write-Log "Updating Room Mailbox $UserName : Adding Permissions"
                    $GroupName = "ro_$UserName"
                    Add-MailboxPermission -Identity $UserName -User $GroupName -AccessRights FullAccess -confirm:$false
                    Add-RecipientPermission -Identity $UserName -Trustee $GroupName -Accessrights "SendAs" -confirm:$false
                    Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                }
                default {
                    if ($Dept -eq "IT" -and $HiPriv -eq "Y") {
                        $UserName = $UserName + "-admin"
                        if ($PrivLevel -ge "1") {
                            foreach ($roleName in $Level1Roles) {
                                Write-Log "Assigning roles for $UserName"
                                Add-MsolRoleMember -RoleMemberEmailAddress "$UserName@$EmailSuffix" -RoleName $roleName
                            }
                        }
                        if ($PrivLevel -ge "2") {
                            foreach ($roleName in $Level2Roles) {
                                Write-Log "Assigning roles for $UserName"
                                Add-MsolRoleMember -RoleMemberEmailAddress "$UserName@$EmailSuffix" -RoleName $roleName
                            }
                        }
                        if ($PrivLevel -ge "3") {
                            foreach ($roleName in $Level3Roles) {
                                Write-Log "Assigning roles for $UserName"
                                Add-MsolRoleMember -RoleMemberEmailAddress "$UserName@$EmailSuffix" -RoleName $roleName
                            }
                            $UserName = "da-" + $UserName
                            foreach ($roleName in $Roles) {
                                Write-Log "Assigning roles for $UserName"
                                Add-MsolRoleMember -RoleMemberEmailAddress "$UserName@$EmailSuffix" -RoleName $roleName
                            }
                        }
                    }
                }
            }
        } else {
            $logmsg = "Mailbox:" + $UserName +" not found in AzureAD"
            Write-Log $logMsg
        }
    }
    Write-Log ""
    Write-Log "Office 365 sync & mailbox update complete"
} catch {
    $e = $_.Exception
    $line = $_.InvocationInfo.ScriptLineNumber
    $msg = $e.Message
    $Action = "Failed to Complete Mailbox Update"
    Write-Log $e $line $msg $Action
    Write-Log $Action
}
#====================================================================
