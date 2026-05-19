#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet("S","H")][string]$UserType
    ,[string]$UserName
    ,[switch]$NonInteractive
    ,[string]$LogFile
)

Set-StrictMode -Version Latest
Add-Type -AssemblyName "microsoft.visualbasic" -ErrorAction Stop

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
# Set variables
$ScriptPath = $PSScriptRoot
$ScriptTitle = "$Domain Password Change Script"
$PasswordLength = $Env.Security.PasswordLength # Number of characters per password
# File locations
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_password_change_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}
# Audit trail: durable record of every password-reset attempt (one row per
# submitted reset, written from a finally block so partial / failed resets
# still get a row).
$AuditFile = Join-Path $ScriptPath "LogFiles\password-resets.csv"

Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-LogFile -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-LogFile -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-LogFile -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-LogFile -LogFile $LogFile -LogString "$ScriptTitle"
Write-LogFile -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "

# Interactive unless we're headless (scheduled task / remoting / Server Core)
# or the caller forces it. Drives whether results are shown as a modal dialog
# or written to the console, so the script never blocks unattended and needs
# no GUI subsystem.
$Interactive = [Environment]::UserInteractive -and -not $NonInteractive

# Surface a result to the operator. Interactive -> the familiar dialog;
# non-interactive -> the same text on the console. NEVER written to the log:
# the success message carries the temporary password, which must not persist.
function Show-OperatorMessage {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Information','Exclamation')][string]$Icon = 'Information'
    )
    if ($Interactive) {
        $null = [microsoft.visualbasic.interaction]::Msgbox($Message, "OKOnly,$Icon", $ScriptTitle)
    } else {
        Write-Host $Message -ForegroundColor $(if ($Icon -eq 'Exclamation') { 'Red' } else { 'Green' })
    }
}

$requiredGroups = @("$($Env.Groups.TaskPrefix)Standard_Account_Admins", 'Domain Admins')
if (-not (Test-IsMemberOf -Sam $env:USERNAME -GroupNames $requiredGroups -DCHostName $DCHostName)) {
    Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}
switch ($UserType.ToUpperInvariant()) {
    "S" {
        $OU = $Env.OUs.Staff
    }
    "H" {
        $requiredGroups = @("$($Env.Groups.TaskPrefix)HiPriv_Account_Admins", 'Domain Admins')
        if (-not (Test-IsMemberOf -Sam $env:USERNAME -GroupNames $requiredGroups -DCHostName $DCHostName)) {
            Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
            throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
        }
        $OU = "$($Env.OUs.HiPrivAccounts),OU=$($Env.OUs.Administration)"
    }
    default {
        throw "Invalid UserType. Use 'S' (Staff) or 'H' (High Privilege)."
    }
}
$OUPath = "OU=$OU,$EndPath"

if (!$UserName) {
    if (-not $Interactive) {
        throw "No -UserName supplied in a non-interactive session: the Out-GridView picker is unavailable (and absent on Server Core). Re-run with -UserName."
    }
    $UserObject = Get-ADUser -filter "enabled -eq 'true'" -SearchBase $OUPath -SearchScope Subtree -Properties Name,SamAccountName,GivenName,Surname,DistinguishedName,Department -Server $DCHostName |
    Out-GridView -title "Select a user account or cancel" -OutputMode Single
} else {
    $UserObject = Get-ADUser -filter 'SamAccountName -eq $UserName' -SearchBase $OUPath -SearchScope Subtree -Properties Name,SamAccountName,GivenName,Surname,DistinguishedName,Department -Server $DCHostName
    if (!$UserObject) {
        Write-LogFile -LogFile $LogFile -LogString "REFUSED: '$UserName' was not found as an account under '$OUPath' (UserType '$UserType'). No reset attempted."
        throw "User '$UserName' was not found as an account under '$OUPath'. Refusing to reset an account outside the '$UserType' scope."
    }
}
if ($UserObject) {
    #generate the new password
    $Password = New-Password -Length $PasswordLength
    #only continue is there is text for the password
    if (-not [string]::IsNullOrWhiteSpace($Password)) {
        # Audit setup: capture the attempt time before any work happens, and
        # default to "Failed" so any early exit / uncaught path audits as
        # such. Specific paths overwrite this to "Success" or "PolicyViolation".
        $ResetTime = Get-Date
        $Outcome = "Failed"
        try {
            #Validate Password against Password Policy
            try {
                Test-Password -LogFile $LogFile -Password $Password -PasswordLength $PasswordLength
            } catch {
                $Outcome = "PolicyViolation"
                Show-OperatorMessage -Message "Password does not meet policy requirements. Please try again." -Icon Exclamation
                return
            }
            #convert to secure string
            $NewPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
            # NOTE: $Password is intentionally retained until the success dialog
            # so the operator can read the temp password. It is cleared (along
            # with $NewPassword) in the finally block, on every exit path.
            # define a hash table of parameter values to splat to Set-ADAccountPassword
            $paramHash = @{
                Identity = $UserObject.SamAccountName
                NewPassword = $NewPassword
                Reset = $True
                Passthru = $True
                ErrorAction = "Stop"
            }
            try {
                $output = Set-ADAccountPassword @paramHash -Server $DCHostName |
                Set-ADUser -ChangePasswordAtLogon $True -PassThru -Server $DCHostName |
                Get-ADuser -Properties PasswordLastSet,PasswordExpired,WhenChanged -Server $DCHostName |
                Out-String -Width 4096
                #display the new temp password + user details in a message box
                $message = "New temporary password: $Password`r`n`r`n$output"
                Show-OperatorMessage -Message $message -Icon Information
                $Outcome = "Success"
            } catch {
                #display error in a message box
                $message = "Failed to reset password for $($UserObject.SamAccountName). $($_.Exception.Message)"
                Show-OperatorMessage -Message $message -Icon Exclamation
                # $Outcome stays at its default "Failed"
            }
        } finally {
            #====================================================================
            # Clear the plaintext and secure-string copies on every exit path
            # (success, failure, policy-violation return).
            #====================================================================
            $Password = $null
            $NewPassword = $null
            #====================================================================
            # Append to audit trail CSV
            #====================================================================
            $AuditRow = [pscustomobject]@{
                Timestamp = $ResetTime.ToString("yyyy-MM-dd HH:mm:ss")
                Operator  = "$Domain\$env:USERNAME"
                Account   = $UserObject.SamAccountName
                UserType  = $UserType
                Outcome   = $Outcome
            }
            try {
                if (Test-Path $AuditFile) {
                    $AuditRow | Export-Csv -Path $AuditFile -NoTypeInformation -Append
                } else {
                    $AuditRow | Export-Csv -Path $AuditFile -NoTypeInformation
                }
            } catch {
                # Don't let a failed audit write mask the reset outcome; record
                # the row (and the failure) to the log file as a fallback.
                Write-LogFile -LogFile $LogFile -LogString "WARNING: failed to write audit row to '$AuditFile'. $($_.Exception.Message)"
                Write-LogFile -LogFile $LogFile -LogString "AUDIT (fallback): $($AuditRow.Timestamp) | $($AuditRow.Operator) | $($AuditRow.Account) | $($AuditRow.UserType) | $($AuditRow.Outcome)"
            }
        }
    } #if plain text password
}
