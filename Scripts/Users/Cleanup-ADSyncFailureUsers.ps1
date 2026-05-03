#Requires -Modules ActiveDirectory, Microsoft.Graph, ExchangeOnlineManagement, ADSync
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
    , [string]$LogFile
)

Set-StrictMode -Version Latest

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

$Domain = "$env:userdomain"
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$PasswordLength = 20 # Number of characters per password
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

Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-Log -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-Log -LogFile $LogFile -LogString "$ScriptTitle"
Write-Log -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString " "
$requiredGroups = @('ADM_Task_Standard_Account_Admins', 'ADM_Task_Standard_Group_Admins', 'ADM_Task_SER_Account_Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

#====================================================================
# Loop through CSV & create users
#====================================================================
# Read input file
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString "Reading user data from input file '$UserInputFile'"
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString " "
# Read list of users from CSV file ignoring first line
$UserList = @(Import-CSV $UserInputFile)
if ($UserList -isnot [Array]) {$UserList = @($UserList)}
$RequiredHeaders = @(
    "USERNAME","FIRSTNAME","LASTNAME","DEPT","COMPANY","S/E/R","CAP","HIPRIV","PrivLevel"
)
$Headers = ($UserList | Select-Object -First 1).PSObject.Properties.Name
foreach ($h in $RequiredHeaders) {
    if ($Headers -notcontains $h) {
        throw "CSV missing required column '$h'"
    }
}
try {
    if (!(Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Log -LogFile $LogFile -LogString "Microsoft.Graph module not installed"
        throw "Microsoft.Graph module not installed"
    }
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log -LogFile $LogFile -LogString "ExchangeOnlineManagement module not installed"
        throw "ExchangeOnlineManagement module not installed"
    }
    Write-Log -LogFile $LogFile -LogString "Starting AzureAD Sync"
    Import-Module -Name ADSync
    Start-ADSyncSyncCycle -PolicyType Delta
    Write-Log -LogFile $LogFile -LogString "Connecting to Microsoft Graph"
    Import-Module -Name Microsoft.Graph.Authentication
    Import-Module -Name Microsoft.Graph.Users
    Import-Module -Name Microsoft.Graph.Identity.Governance
    Connect-MgGraph -NoWelcome -Scopes "RoleManagement.ReadWrite.Directory", "User.ReadWrite.All"
    Write-Log -LogFile $LogFile -LogString "Connected to Microsoft Graph"
    Write-Log -LogFile $LogFile -LogString "Connecting to Exchange Online"
    Import-Module -Name ExchangeOnlineManagement
    Connect-ExchangeOnline
    Write-Log -LogFile $LogFile -LogString "Connected to Exchange Online"
    Write-Log -LogFile $LogFile -LogString "Pausing for 30 seconds"
    Start-Sleep -s 30
    # Process each input file record
    foreach ($USER in $UserList) {
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
            $i = 0
            $MBX = $null
            Do {
                $MBX = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 30
                $i++
                if ($i -eq 4) {
                    Start-ADSyncSyncCycle -PolicyType Delta
                    $i = 0
                }
            } While (!($MBX) -and $i -lt 5)
            if ($MBX) {
                if ($SharedEquipmentRoom) {
                    $logmsg = "Updating Mailbox: " + $UserPrincipalName +" "+ $SharedEquipmentRoom +" "+ $Capacity
                    Write-Log -LogFile $LogFile -LogString $LogMsg
                } elseif (!$SharedEquipmentRoom -and !$Capacity) {
                    $logmsg = "Updating Mailbox: " + $UserPrincipalName
                    Write-Log -LogFile $LogFile -LogString $LogMsg
                }
                Update-UserMailbox -LogFile $LogFile -UserPrincipalName $UserPrincipalName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
            } else {
                $logmsg = "Mailbox: " + $UserPrincipalName +" not found in AzureAD"
                Write-Log -LogFile $LogFile -LogString $LogMsg
            }
            if ($Dept -eq "IT" -and $HiPriv -eq "Y") {
                Write-Log -LogFile $LogFile -LogString "Creating Cloud Admin account for $UserName"
                Write-Log -LogFile $LogFile -LogString " "
                & $PSScriptRoot\CreateITCloudAdminUser.ps1 -FirstName $FirstName -LastName $LastName -UserName $UserName -EmailSuffix $EmailSuffix -PrivLevel $PrivLevel -Dept $Dept -Company $Company -LogFile $LogFile -Manager $UserName -PasswordLength $PasswordLength
            }
        } catch {
            $e = $_.Exception
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
            Write-Log -LogFile $LogFile -LogString "ERROR processing '$UserName' - Line $line : $($e.Message)" -ForegroundColor Red
            Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
            continue
        }
    }
    Write-Log -LogFile $LogFile -LogString " "
    Write-Log -LogFile $LogFile -LogString "Office 365 sync & mailbox update complete"
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
} catch {
    $e = $_.Exception
    $line = $_.InvocationInfo.ScriptLineNumber
    $msg = $e.Message
    $Action = "Failed to Complete Mailbox Update"
    Write-Log -LogFile $LogFile -LogString "Exception: $($e.Message)"
    Write-Log -LogFile $LogFile -LogString "Line: $line"
    Write-Log -LogFile $LogFile -LogString $msg
    Write-Log -LogFile $LogFile -LogString $Action
}
#====================================================================
