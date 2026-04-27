#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
)

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
