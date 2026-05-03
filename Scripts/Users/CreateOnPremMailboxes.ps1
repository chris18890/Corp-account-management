#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
)

Set-StrictMode -Version Latest

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

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
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-Log -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-Log -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-Log -LogFile $LogFile -LogString "$ScriptTitle"
Write-Log -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString " "
$requiredGroups = @('ADM_Task_Standard_Account_Admins', 'ADM_Task_SER_Account_Admins', 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

# Get user credentials for server connectivity (Non-MFA)
try {
    $Cred = Get-Credential -ErrorAction Stop -Message "Admin credentials for remote sessions:"
} catch {
    $ErrorMsg = $_.Exception.Message
    Write-Log -LogFile $LogFile -LogString "Failed to validate credentials: $ErrorMsg "
    Read-Host -Prompt "Press Enter to exit"
    Exit
}
#Connect to remote Exchange PowerShell
Write-Log -LogFile $LogFile -LogString "Connecting to remote Exchange PowerShell session... "
try {
    $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://$ExServer/PowerShell" -Name ExSession -Credential $cred -ErrorAction Stop -authentication Kerberos
    Write-Log -LogFile $LogFile -LogString "connected."
    Write-Log -LogFile $LogFile -LogString "Importing Exchange session... "
    Import-PSSession -Session $ExSession -ErrorAction Stop -AllowClobber > $null
    Write-Log -LogFile $LogFile -LogString "done."
} catch {
    $e = $_.Exception
    Write-Log -LogFile $LogFile -LogString $e
    $line = $_.InvocationInfo.ScriptLineNumber
    Write-Log -LogFile $LogFile -LogString $line
    $msg = $e.Message
    Write-Log -LogFile $LogFile -LogString $msg
    $Action = "Error Importing Exchange Session"
    Write-Log -LogFile $LogFile -LogString $Action
    Write-Log -LogFile $LogFile -LogString "failed."
    Write-Log -LogFile $LogFile -LogString "ERROR: $_" -ForegroundColor Red
}
if (!$ExSession) {
    Write-Log -LogFile $LogFile -LogString "Exchange session not connected Stopping Script"
    Exit
}
#====================================================================

#====================================================================
#Loop through CSV & create users
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
    "USERNAME","S/E/R","CAP","REALNAME"
)
$Headers = ($UserList | Select-Object -First 1).PSObject.Properties.Name
foreach ($h in $RequiredHeaders) {
    if ($Headers -notcontains $h) {
        throw "CSV missing required column '$h'"
    }
}
# Process each input file record
foreach ($USER in $UserList) {
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
            Write-Log -LogFile $LogFile -LogString ("=" * 80)
            Write-Log -LogFile $LogFile -LogString "Processing input file record for $UserName..."
            Write-Log -LogFile $LogFile -LogString ("=" * 80)
            Write-Log -LogFile $LogFile -LogString "Exchange mailbox for $UserName will be created in Exchange OnPrem"
            Write-Log -LogFile $LogFile -LogString "Calling New-UserOnPremMailbox function with the following parameters:"
            Write-Log -LogFile $LogFile -LogString "UserName: $UserName, EmailSuffix $EmailSuffix, realname: $RealName, SharedEquipmentRoom: $SharedEquipmentRoom, Capacity: $Capacity"
            $EnabledMailboxes += New-UserOnPremMailbox -LogFile $LogFile -DCHostName $DCHostName -UserName $UserName -EmailSuffix $EmailSuffix -realname $RealName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
            Write-Log -LogFile $LogFile -LogString ("=" * 80)
            Write-Log -LogFile $LogFile -LogString "Processing input file record for $UserName complete"
            Write-Log -LogFile $LogFile -LogString ("=" * 80)
            Write-Log -LogFile $LogFile -LogString " "
        } catch {
            $e = $_.Exception
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
            Write-Log -LogFile $LogFile -LogString "ERROR: Failed during processing of $UserName - Line $Line" -ForegroundColor Red
            Write-Log -LogFile $LogFile -LogString "$e"
            Write-Log -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
            Write-Log -LogFile $LogFile -LogString ("=" * 80)
            Write-Log -LogFile $LogFile -LogString " "
        }
    }
}
Write-Log -LogFile $LogFile -LogString "Updating Mailboxes"
foreach ($mailbox in $EnabledMailboxes) {
    $i = 0
    $MBX = $null
    Do {
        $MBX = Get-Mailbox -Identity $mailbox.alias -DomainController $DCHostName -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 30
        $i++
    } While (!($MBX) -and $i -lt 5)
    if ($MBX) {
        if ($Mailbox.SharedEquipmentRoom) {
            $logmsg = "Updating Mailbox: " + $Mailbox.Alias +" "+ $Mailbox.SharedEquipmentRoom +" "+ $Mailbox.Capacity +" "
            Write-Log -LogFile $LogFile -LogString $LogMsg
        } elseif (!$Mailbox.SharedEquipmentRoom -and !$Mailbox.Capacity) {
            $logmsg = "Updating Mailbox: " + $Mailbox.Alias
            Write-Log -LogFile $LogFile -LogString $LogMsg
        }
        Update-UserOnPremMailbox -LogFile $LogFile -DCHostName $DCHostName -UserName $mailbox.Alias -SharedEquipmentRoom $Mailbox.SharedEquipmentRoom -Capacity $Mailbox.Capacity
    } else {
        $logmsg = "Mailbox: " + $Mailbox.Alias +" not found in AD"
        Write-Log -LogFile $LogFile -LogString $LogMsg
    }
}
if (Get-PSSession -Name ExSession -ErrorAction SilentlyContinue) {
    Remove-PsSession $ExSession
    $Cred.Password.Dispose()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Log -LogFile $LogFile -LogString "Closed Exchange session."
}
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString "Processing complete"
Write-Log -LogFile $LogFile -LogString ("=" * 80)
#====================================================================
