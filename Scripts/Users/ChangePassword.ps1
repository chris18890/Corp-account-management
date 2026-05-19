#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet("S","H")][string]$UserType
    ,[string]$LogFile
)

Set-StrictMode -Version Latest
Add-Type -AssemblyName "microsoft.visualbasic" -ErrorAction Stop

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
switch ($UserType.ToUpperInvariant()) {
    "S" {
        $OU = $Env.OUs.Staff
    }
    "H" {
        $OU = "$($Env.OUs.HiPrivAccounts),OU=$($Env.OUs.Administration)"
    }
    default {
        throw "Invalid UserType. Use 'S' (Staff) or 'H' (High Privilege)."
    }
}
$OUPath = "OU=$OU,$EndPath"
$ScriptPath = $PSScriptRoot
# Set variables
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

#get all enabled user accounts in the OU
$User = Get-ADUser -filter "enabled -eq 'true'" -SearchBase $OUPath -SearchScope Subtree -Properties Name,SamAccountName,GivenName,Surname,DistinguishedName,Department -Server $DCHostName |
Out-GridView -title "Select a user account or cancel" -OutputMode Single
if ($User) {
    #prompt for the new password
    $Password = (Get-Credential -Message "Please enter new password for $($User.SamAccountName)" -User "$Domain\$($User.SamAccountName)").getNetworkCredential().password
    #only continue is there is text for the password
    if (-not [string]::IsNullOrWhiteSpace($Password)) {
        # Audit setup: capture the attempt time before any work happens, and
        # default to "Failed" so any early exit / uncaught path audits as
        # such. Specific paths overwrite this to "Success" or "PolicyViolation".
        $ResetTime = Get-Date
        $Outcome   = "Failed"
        try {
            #Validate Password against Password Policy
            try {
                Test-Password -LogFile $LogFile -Password $Password -PasswordLength $PasswordLength
            } catch {
                $Outcome = "PolicyViolation"
                [microsoft.visualbasic.interaction]::Msgbox(
                    "Password does not meet policy requirements. Please try again.",
                    "OKOnly,Exclamation",
                    $ScriptTitle
                ) | Out-Null
                return
            }
            #convert to secure string
            $NewPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
            $Password = $null
            #define a hash table of parameter values to splat to
            #Set-ADAccountPassword
            $paramHash = @{
                Identity = $User.SamAccountName
                NewPassword = $NewPassword
                Reset = $True
                Passthru = $True
                ErrorAction = "Stop"
            }
            try {
                $output = Set-ADAccountPassword @paramHash -Server $DCHostName |
                Set-ADUser -ChangePasswordAtLogon $True -PassThru -Server $DCHostName |
                Get-ADuser -Properties PasswordLastSet,PasswordExpired,WhenChanged -Server $DCHostName |
                Out-String
                #display user in a message box
                $message = $output
                $button = "OKOnly"
                $icon = "Information"
                [microsoft.visualbasic.interaction]::Msgbox($message,"$button,$icon",$ScriptTitle) | Out-Null
                $Outcome = "Success"
            } catch {
                #display error in a message box
                $message =  "Failed to reset password for $($User.SamAccountName). $($_.Exception.Message)"
                $button = "OKOnly"
                $icon = "Exclamation"
                [microsoft.visualbasic.interaction]::Msgbox($message,"$button,$icon",$ScriptTitle) | Out-Null
                # $Outcome stays at its default "Failed"
            }
            $NewPassword = $null
        } finally {
            #====================================================================
            # Append to audit trail CSV
            #====================================================================
            $AuditRow = [pscustomobject]@{
                Timestamp = $ResetTime.ToString("yyyy-MM-dd HH:mm:ss")
                Operator  = "$Domain\$env:USERNAME"
                Account   = $User.SamAccountName
                UserType  = $UserType
                Outcome   = $Outcome
            }
            if (Test-Path $AuditFile) {
                $AuditRow | Export-Csv -Path $AuditFile -NoTypeInformation -Append
            } else {
                $AuditRow | Export-Csv -Path $AuditFile -NoTypeInformation
            }
        }
    } #if plain text password
}
