#Requires -Modules ActiveDirectory, Microsoft.Graph, ExchangeOnlineManagement, ADSync
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
    , [string]$O365EmailSuffix
    , [string]$LogFile
)

Set-StrictMode -Version Latest

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

$Domain = "$env:userdomain"
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
if (!$O365EmailSuffix) {
    $O365EmailSuffix = READ-HOST 'Enter "onmicrosoft.com" domain - '
}
if ($O365EmailSuffix -notmatch '\.onmicrosoft\.com$') {
    $O365EmailSuffix = "$O365EmailSuffix.onmicrosoft.com"
}
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$AzureADConnect = "$Domain-RTR.$DNSSuffix"
$PasswordLength = $Env.Security.PasswordLength # Number of characters per password
$ScriptPath = $PSScriptRoot
$ScriptTitle = "$Domain User Sync Script"
$LogPath = "$ScriptPath\LogFiles"
$UserInputFile = "$ScriptPath\users.csv"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_new_user_cleanup_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-LogFile -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-LogFile -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-LogFile -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-LogFile -LogFile $LogFile -LogString "$ScriptTitle"
Write-LogFile -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "

$requiredGroups = @("$($Env.Groups.TaskPrefix)Standard_Account_Admins", "$($Env.Groups.TaskPrefix)Standard_Group_Admins", "$($Env.Groups.TaskPrefix)SER_Account_Admins", 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}
try {
    $Cred = Get-Credential -ErrorAction Stop -Message "Admin credentials for remote sessions:"
} catch {
    $ErrorMsg = $_.Exception.Message
    Write-LogFile -LogFile $LogFile -LogString "Failed to validate credentials: $ErrorMsg "
    Read-Host -Prompt "Press Enter to exit"
    Exit
}

#====================================================================
# Loop through CSV & create users
#====================================================================
# Read input file
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Reading user data from input file '$UserInputFile'"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "
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
        Write-LogFile -LogFile $LogFile -LogString "Microsoft.Graph module not installed"
        throw "Microsoft.Graph module not installed"
    }
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-LogFile -LogFile $LogFile -LogString "ExchangeOnlineManagement module not installed"
        throw "ExchangeOnlineManagement module not installed"
    }
    Write-LogFile -LogFile $LogFile -LogString "Starting AzureAD Sync"
    Import-Module -Name ADSync
    Invoke-ADSync -LogFile $LogFile -Cred $Cred -AzureADConnect $AzureADConnect -O365EmailSuffix $O365EmailSuffix
    Write-LogFile -LogFile $LogFile -LogString "Connecting to Exchange Online"
    Import-Module -Name ExchangeOnlineManagement
    Connect-ExchangeOnline
    Write-LogFile -LogFile $LogFile -LogString "Connected to Exchange Online"
    Write-LogFile -LogFile $LogFile -LogString "Connecting to Microsoft Graph"
    Import-Module -Name Microsoft.Graph.Authentication
    Import-Module -Name Microsoft.Graph.Users
    Import-Module -Name Microsoft.Graph.Identity.Governance
    Connect-MgGraph -NoWelcome -Scopes "RoleManagement.ReadWrite.Directory", "User.ReadWrite.All"
    Write-LogFile -LogFile $LogFile -LogString "Connected to Microsoft Graph"
    Write-LogFile -LogFile $LogFile -LogString "Pausing for 30 seconds"
    Start-Sleep -s 30
    # Process each input file record
    foreach ($USER in $UserList) {
        try {
            $FirstName = ConvertTo-SafeName $USER.FIRSTNAME
            $LastName = ConvertTo-SafeName $USER.LASTNAME
            $UserName = ConvertTo-SafeSamAccountName $USER.USERNAME
            $Company = $USER.COMPANY
            $Dept = $USER.DEPT
            $HiPriv = $USER.HIPRIV.ToUpper()
            [int]$PrivLevel = ConvertTo-IntOrDefault $USER.PrivLevel
            $SharedEquipmentRoom = $USER.'S/E/R'.ToUpper()
            [int]$Capacity = ConvertTo-IntOrDefault $USER.Cap
            $UserPrincipalName = "$UserName@$EmailSuffix"
            $i = 0
            $MBX = $null
            Do {
                $MBX = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 30
                $i++
                if ($i -eq 4) {
                    Invoke-ADSync -LogFile $LogFile -Cred $Cred -AzureADConnect $AzureADConnect -O365EmailSuffix $O365EmailSuffix
                    $i = 0
                }
            } While (!($MBX) -and $i -lt 5)
            if ($MBX) {
                if ($SharedEquipmentRoom) {
                    $logmsg = "Updating Mailbox: $UserPrincipalName $SharedEquipmentRoom $Capacity"
                    Write-LogFile -LogFile $LogFile -LogString $LogMsg
                } elseif (!$SharedEquipmentRoom -and !$Capacity) {
                    $logmsg = "Updating Mailbox: $UserPrincipalName"
                    Write-LogFile -LogFile $LogFile -LogString $LogMsg
                }
                Update-UserMailbox -LogFile $LogFile -UserPrincipalName $UserPrincipalName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
            } else {
                $logmsg = "Mailbox: $UserPrincipalName not found in AzureAD"
                Write-LogFile -LogFile $LogFile -LogString $LogMsg
            }
            if ($Dept -eq "IT" -and $HiPriv -eq "Y") {
                Write-LogFile -LogFile $LogFile -LogString "Creating Cloud Admin account for $UserName"
                Write-LogFile -LogFile $LogFile -LogString " "
                & $PSScriptRoot\CreateITCloudAdminUser.ps1 -FirstName $FirstName -LastName $LastName -UserName $UserName -EmailSuffix $EmailSuffix -PrivLevel $PrivLevel -Dept $Dept -Company $Company -LogFile $LogFile -Manager $UserName -PasswordLength $PasswordLength
            }
        } catch {
            $e = $_.Exception
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-LogFile -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
            Write-LogFile -LogFile $LogFile -LogString "ERROR processing '$UserName' - Line $line : $($e.Message)" -ForegroundColor Red
            Write-LogFile -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
            continue
        }
    }
    Write-LogFile -LogFile $LogFile -LogString " "
    Write-LogFile -LogFile $LogFile -LogString "Office 365 sync & mailbox update complete"
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
    if (Get-PSSession) {
        Write-LogFile -LogFile $LogFile -LogString "Cleaning up PSSessions"
        Get-PSSession | Remove-PSSession
        $Cred.Password.Dispose()
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
} catch {
    $e = $_.Exception
    $line = $_.InvocationInfo.ScriptLineNumber
    $msg = $e.Message
    $Action = "Failed to Complete Mailbox Update"
    Write-LogFile -LogFile $LogFile -LogString "Exception: $($e.Message)"
    Write-LogFile -LogFile $LogFile -LogString "Line: $line"
    Write-LogFile -LogFile $LogFile -LogString $msg
    Write-LogFile -LogFile $LogFile -LogString $Action
} finally {
    if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    if (Get-PSSession) {
        Get-PSSession | Remove-PSSession
        $Cred.Password.Dispose()
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
#====================================================================
