#Requires -Modules ActiveDirectory, Microsoft.Graph, ExchangeOnlineManagement, ADSync
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
)

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

$Domain = "$env:userdomain"
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptTitle = "$Domain User Sync Script"
$LogPath = "$ScriptPath\LogFiles"
$UserInputFile = "$ScriptPath\users.csv"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
$LogFileName = $Domain + "_new_user_cleanup_log-$(Get-Date -Format 'yyyyMMdd')"
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
$requiredGroups = @('ADM_Task_Standard_Account_Admins', 'ADM_Task_Standard_Group_Admins', 'ADM_Task_SER_Account_Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

#====================================================================
# Loop through CSV & create users
#====================================================================
# Read input file
Write-Log ("=" * 80)
Write-Log "Reading user data from input file '$UserInputFile'"
Write-Log ("=" * 80)
Write-Log ""
# Read list of users from CSV file ignoring first line
$CreatedUsers = @(Import-CSV $UserInputFile)
$RequiredHeaders = @(
    "USERNAME","FIRSTNAME","LASTNAME","DEPT","COMPANY","S/E/R","CAP","HIPRIV","PrivLevel"
)
$Headers = ($CreatedUsers | Select-Object -First 1).PSObject.Properties.Name
foreach ($h in $RequiredHeaders) {
    if ($Headers -notcontains $h) {
        throw "CSV missing required column '$h'"
    }
}
try {
    if (!(Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Log "Microsoft.Graph module not installed"
        throw "Microsoft.Graph module not installed"
    }
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log "ExchangeOnlineManagement module not installed"
        throw "ExchangeOnlineManagement module not installed"
    }
    Write-Log "Starting AzureAD Sync"
    Import-Module -Name ADSync
    Start-ADSyncSyncCycle -PolicyType Delta
    Write-Log "Connecting to Microsoft Graph"
    Import-Module -Name Microsoft.Graph.Authentication
    Import-Module -Name Microsoft.Graph.Users
    Import-Module -Name Microsoft.Graph.Identity.Governance
    Connect-MgGraph -NoWelcome -Scopes "RoleManagement.ReadWrite.Directory", "User.ReadWrite.All"
    Write-Log "Connected to Microsoft Graph"
    Write-Log "Connecting to Exchange Online"
    Import-Module -Name ExchangeOnlineManagement
    Connect-ExchangeOnline
    Write-Log "Connected to Exchange Online"
    Write-Log "Pausing for 30 seconds"
    Start-Sleep -s 30
    # Process each input file record
    foreach ($USER in $CreatedUsers) {
        try {
            $FirstName = $USER.FIRSTNAME
            $FirstName = $FirstName.Trim() -replace '[?@\\+]', [String]::Empty # Strip out illegal characters from First name for Office 365 compliance. Note that \ is escaped to \\
            $LastName = $USER.LASTNAME
            $LastName = $LastName.Trim() -replace '[?@\\+]', [String]::Empty # Strip out illegal characters from LastName for Office 365 compliance. Note that \ is escaped to \\
            $UserName = $USER.USERNAME
            $UserName = $UserName.Trim() -replace '[^A-Za-z0-9.-]', [String]::Empty # Strip out illegal characters from User ID
            if ($UserName.Length -gt 20) {
                $UserName = $UserName.Substring(0,20)
            }
            $Company = $USER.COMPANY
            $Dept = $USER.DEPT
            $HiPriv = $USER.HIPRIV.ToUpper()
            [int]$PrivLevel = $USER.PrivLevel
            $SharedEquipmentRoom = $USER.'S/E/R'.ToUpper()
            [int]$Capacity = $USER.Cap
            $UserPrincipalName = "$UserName@$EmailSuffix"
            $MBX = $null
            $MBX = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
            $i = 0
            while (!($MBX) -and ($i -le 6)) {
                $MBX = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
                $i++
                Start-Sleep -seconds 10
            }
            if ($MBX) {
                Write-Log ""
                Write-Log "Assigning region for $UserName"
                Update-MgUser -UserId $UserPrincipalName -UsageLocation GB
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
                        $GroupName = "sh_$UserName@$EmailSuffix"
                        Add-MailboxPermission -Identity $UserPrincipalName -User $GroupName -AccessRights FullAccess -confirm:$false
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
                        Write-Log "Updating Equipment Mailbox $UserName : Updating Calendar Processing"
                        Set-CalendarProcessing -Identity $UserPrincipalName -AutomateProcessing AutoAccept -confirm:$false
                        if ($Capacity) {
                            Set-Mailbox -Identity $UserPrincipalName -ResourceCapacity $Capacity
                        }
                        Write-Log "Updating Equipment Mailbox $UserName : Adding Permissions"
                        $GroupName = "eq_$UserName@$EmailSuffix"
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
                        Write-Log "Updating Room Mailbox $UserName : Updating Calendar Processing"
                        Set-CalendarProcessing -Identity $UserPrincipalName -AutomateProcessing AutoAccept -confirm:$false
                        if ($Capacity) {
                            Set-Mailbox -Identity $UserPrincipalName -ResourceCapacity $Capacity
                        }
                        Write-Log "Updating Room Mailbox $UserName : Adding Permissions"
                        $GroupName = "ro_$UserName@$EmailSuffix"
                        Add-MailboxPermission -Identity $UserPrincipalName -User $GroupName -AccessRights FullAccess -confirm:$false
                        Add-RecipientPermission -Identity $UserPrincipalName -Trustee $GroupName -Accessrights "SendAs" -confirm:$false
                        Write-Log "Delegated permissions for mailbox $UserName to group $GroupName"
                    }
                    default {
                        if ($Dept -eq "IT" -and $HiPriv -eq "Y") {
                            Write-Log "Creating Cloud Admin account for $UserName"
                            Write-Log ""
                            .\CreateITCloudAdminUser.ps1 -FirstName $FirstName -LastName $LastName -UserName $UserName -EmailSuffix $EmailSuffix -PrivLevel $PrivLevel -Dept $Dept -Company $Company -LogFile $LogFile -Manager $UserName
                        }
                    }
                }
            } else {
                $logmsg = "Mailbox: " + $UserName +" not found in AzureAD"
                Write-Log $logMsg
            }
        } catch {
            $e = $_.Exception
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log ("-" * 80) -ForegroundColor Red
            Write-Log "ERROR processing '$UserName' - Line $line : $($e.Message)" -ForegroundColor Red
            Write-Log ("-" * 80) -ForegroundColor Red
            continue
        }
    }
    Write-Log ""
    Write-Log "Office 365 sync & mailbox update complete"
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
} catch {
    $e = $_.Exception
    $line = $_.InvocationInfo.ScriptLineNumber
    $msg = $e.Message
    $Action = "Failed to Complete Mailbox Update"
    Write-Log "Exception: $($e.Message)"
    Write-Log "Line: $line"
    Write-Log $msg
    Write-Log $Action
}
#====================================================================
