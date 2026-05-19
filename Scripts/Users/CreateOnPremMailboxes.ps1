#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
    ,[string]$LogFile
)

Set-StrictMode -Version Latest

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

#====================================================================
# Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
# ADConnect & Exchange settings
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ExServer = "$Domain-EXCH.$DNSSuffix" #Remote Exchange PS session
# Get containing folder for script to locate supporting files
$ScriptPath = $PSScriptRoot
# Set variables
$ScriptTitle = "$Domain User Mailbox Creation Script"
$EnabledMailboxes = @() # Array to Store Completed Mailbox requests for later enumeration
# File locations
$LogPath = "$ScriptPath\LogFiles"
$UserInputFile = "$ScriptPath\users.csv"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_mailbox_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}
#====================================================================

#====================================================================
# Start of script
#====================================================================
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-LogFile -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-LogFile -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-LogFile -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-LogFile -LogFile $LogFile -LogString "$ScriptTitle"
Write-LogFile -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "
$requiredGroups = @("$($Env.Groups.TaskPrefix)Standard_Account_Admins", "$($Env.Groups.TaskPrefix)SER_Account_Admins", 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

# Get user credentials for server connectivity (Non-MFA)
try {
    $Cred = Get-Credential -ErrorAction Stop -Message "Admin credentials for remote sessions:"
} catch {
    $ErrorMsg = $_.Exception.Message
    Write-LogFile -LogFile $LogFile -LogString "Failed to validate credentials: $ErrorMsg "
    Read-Host -Prompt "Press Enter to exit"
    Exit
}
#Connect to remote Exchange PowerShell
Write-LogFile -LogFile $LogFile -LogString "Connecting to remote Exchange PowerShell session... "
try {
    $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://$ExServer/PowerShell" -Name ExSession -Credential $cred -ErrorAction Stop -authentication Kerberos
    Write-LogFile -LogFile $LogFile -LogString "connected."
    Write-LogFile -LogFile $LogFile -LogString "Importing Exchange session... "
    Import-PSSession -Session $ExSession -ErrorAction Stop -AllowClobber > $null
    Write-LogFile -LogFile $LogFile -LogString "done."
} catch {
    $e = $_.Exception
    Write-LogFile -LogFile $LogFile -LogString $e
    $line = $_.InvocationInfo.ScriptLineNumber
    Write-LogFile -LogFile $LogFile -LogString $line
    $msg = $e.Message
    Write-LogFile -LogFile $LogFile -LogString $msg
    $Action = "Error Importing Exchange Session"
    Write-LogFile -LogFile $LogFile -LogString $Action
    Write-LogFile -LogFile $LogFile -LogString "failed."
    Write-LogFile -LogFile $LogFile -LogString "ERROR: $_" -ForegroundColor Red
}
if (!$ExSession) {
    Write-LogFile -LogFile $LogFile -LogString "Exchange session not connected Stopping Script"
    Exit
}
#====================================================================

#====================================================================
#Loop through CSV & create users
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
    $UserName = ConvertTo-SafeSamAccountName $USER.USERNAME
    $SharedEquipmentRoom = $USER.'S/E/R'.ToUpper()
    [int]$Capacity = ConvertTo-IntOrDefault $USER.Cap
    $RealName = $USER.REALNAME
    $ExistingUser = Get-ADObject -LDAPFilter "(&(objectCategory=user)(objectClass=user)(sAMAccountName=$UserName))" -Server $DCHostName
    if ($ExistingUser) {
        try {
            Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
            Write-LogFile -LogFile $LogFile -LogString "Processing input file record for $UserName..."
            Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
            Write-LogFile -LogFile $LogFile -LogString "Exchange mailbox for $UserName will be created in Exchange OnPrem"
            Write-LogFile -LogFile $LogFile -LogString "Calling New-UserOnPremMailbox function with the following parameters:"
            Write-LogFile -LogFile $LogFile -LogString "UserName: $UserName, EmailSuffix $EmailSuffix, realname: $RealName, SharedEquipmentRoom: $SharedEquipmentRoom, Capacity: $Capacity"
            $EnabledMailboxes += New-UserOnPremMailbox -LogFile $LogFile -DCHostName $DCHostName -UserName $UserName -EmailSuffix $EmailSuffix -realname $RealName -SharedEquipmentRoom $SharedEquipmentRoom -Capacity $Capacity
            Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
            Write-LogFile -LogFile $LogFile -LogString "Processing input file record for $UserName complete"
            Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
            Write-LogFile -LogFile $LogFile -LogString " "
        } catch {
            $e = $_.Exception
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-LogFile -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
            Write-LogFile -LogFile $LogFile -LogString "ERROR: Failed during processing of $UserName - Line $Line" -ForegroundColor Red
            Write-LogFile -LogFile $LogFile -LogString "$e"
            Write-LogFile -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
            Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
            Write-LogFile -LogFile $LogFile -LogString " "
        }
    }
}
Write-LogFile -LogFile $LogFile -LogString "Updating Mailboxes"
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
            Write-LogFile -LogFile $LogFile -LogString $LogMsg
        } elseif (!$Mailbox.SharedEquipmentRoom -and !$Mailbox.Capacity) {
            $logmsg = "Updating Mailbox: " + $Mailbox.Alias
            Write-LogFile -LogFile $LogFile -LogString $LogMsg
        }
        Update-UserOnPremMailbox -LogFile $LogFile -DCHostName $DCHostName -UserName $mailbox.Alias -SharedEquipmentRoom $Mailbox.SharedEquipmentRoom -Capacity $Mailbox.Capacity
    } else {
        $logmsg = "Mailbox: " + $Mailbox.Alias +" not found in AD"
        Write-LogFile -LogFile $LogFile -LogString $LogMsg
    }
}
if (Get-PSSession -Name ExSession -ErrorAction SilentlyContinue) {
    Remove-PsSession $ExSession
    $Cred.Password.Dispose()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-LogFile -LogFile $LogFile -LogString "Closed Exchange session."
}
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Processing complete"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
#====================================================================
