#Requires -RunAsAdministrator

[CmdletBinding()]
param([Parameter(Mandatory)][string]$Drive)

Set-StrictMode -Version Latest
Add-Type -AssemblyName Microsoft.UpdateServices.Administration

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

#=========================================
#Domain Names in ADS & DNS format, and main OU name
#=========================================
$Domain = "$env:userdomain"
$ServerName = "$env:computername"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$ParentOU = "Domain Computers"
$Location = "OU=$ParentOU,$EndPath"
$GPOLocation = Join-Path $PSScriptRoot "GPOs"
#=========================================

#=========================================
#Drive where all the folders will be created
#=========================================
$Drive = $Drive.TrimEnd(':') + ':'
$RootShare = "WSUS"
$Feature = "UpdateServices"
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptTitle = "$Domain network setup Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
$LogFileName = $Domain + "_WSUS_setup_log-$(Get-Date -Format 'yyyyMMdd')"
$LogIndex = 0
while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
    $LogIndex ++
}
$LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
#=========================================

Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-Log -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-Log -LogFile $LogFile -LogString "$ScriptTitle"
Write-Log -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString " "

if (!(TEST-PATH "$Drive\$RootShare")) {
    New-Item -Path $Drive -Name $RootShare -ItemType Directory -force
}

if (Get-WindowsFeature -Name $Feature | Where InstallState -Eq Installed) {
    Write-Log -LogFile $LogFile -LogString "$Feature is already installed" -ForegroundColor Green
} else {
    Write-Log -LogFile $LogFile -LogString "installing $Feature"
    install-windowsfeature -IncludeManagementTools -Name $Feature
    Write-Log -LogFile $LogFile -LogString "installed $Feature"
}
cd "C:\Program Files\Update Services\Tools\"
.\wsusutil.exe postinstall CONTENT_DIR=$Drive\$RootShare
if ($LASTEXITCODE -ne 0) { throw "wsusutil postinstall failed - exit code $LASTEXITCODE" }
Import-GPO -BackupGpoName $RootShare -TargetName $RootShare -path $GPOLocation -CreateIfNeeded
try {
    New-GPLink -name $RootShare -target $EndPath -LinkEnabled Yes -enforced yes
    Write-Log -LogFile $LogFile -LogString "Linked $RootShare to $EndPath"
} catch {
    $ex = $_.Exception
    if ($ex.Message -match "already exists") {
        Write-Log -LogFile $LogFile -LogString "'$RootShare' already linked to $EndPath" -ForegroundColor Green
    } else {
        throw
    }
}
Set-GPPermission -Name $RootShare -PermissionLevel GpoEditDeleteModifySecurity -TargetName "ADM_Task_AD_GPO_Admins" -TargetType Group -Server $DCHostName
# Get WSUS Server Object
$wsus = Get-WSUSServer
# Connect to WSUS server configuration
$wsusConfig = $wsus.GetConfiguration()
# Set to download updates from Microsoft Updates
Set-WsusServerSynchronization -SyncFromMU
# Set Update Languages to English and save configuration settings
$wsusConfig.AllUpdateLanguagesEnabled = $false
$wsusConfig.SetEnabledUpdateLanguages("en")
$wsusConfig.Save()
# Get WSUS Subscription and perform initial synchronization to get latest categories
$subscription = $wsus.GetSubscription()
$subscription.StartSynchronizationForCategoryOnly()
write-host "Beginning first WSUS Sync to get available Products etc" -ForegroundColor Magenta
write-host "Will take some time to complete"
While ($subscription.GetSynchronizationStatus() -ne "NotProcessing") {
    Write-Log -LogFile $LogFile -LogString "." -NoNewline
    Start-Sleep -Seconds 5
}
write-host " "
Write-Log -LogFile $LogFile -LogString "Sync is done." -ForegroundColor Green
# Configure the Platforms that we want WSUS to receive updates
write-host "Setting WSUS Products"
Get-WsusProduct | where-Object {
    $_.Product.Title -in (
        "Windows 10",
        "Windows 11",
        "Windows Server 2022",
        "Windows Server 2025"
    )
} | Set-WsusProduct
# Configure the Classifications
write-host "Setting WSUS Classifications"
Get-WsusClassification | Where-Object {
    $_.Classification.Title -in (
        "Critical Updates",
        "Definition Updates",
        "Feature Packs",
        "Security Updates",
        "Service Packs",
        "Update Rollups",
        "Updates"
    )
} | Set-WsusClassification
# Configure Synchronizations
write-host "Enabling WSUS Automatic Synchronisation"
$subscription.SynchronizeAutomatically=$true
# Set synchronization scheduled for midnight each night
$subscription.SynchronizeAutomaticallyTimeOfDay= (New-TimeSpan -Hours 0)
$subscription.NumberOfSynchronizationsPerDay=1
$subscription.Save()
# Kick off a synchronization
$subscription.StartSynchronization()
# Configure Default Approval Rule
write-host "Configuring default automatic approval rule"
$rule = $wsus.GetInstallApprovalRules() | Where-Object {
    $_.Name -eq "Default Automatic Approval Rule"
}
$class = $wsus.GetUpdateClassifications() | Where-Object {
    $_.Title -In (
        "Critical Updates",
        "Definition Updates",
        "Feature Packs",
        "Security Updates",
        "Service Packs",
        "Update Rollups",
        "Updates"
    )
}
$class_coll = New-Object Microsoft.UpdateServices.Administration.UpdateClassificationCollection
$class_coll.AddRange($class)
$rule.SetUpdateClassifications($class_coll)
$rule.Enabled = $True
$rule.Save()

# Run Default Approval Rule
write-host "Running Default Approval Rule"
write-host " >This step may timeout, but the rule will be applied and the script will continue" -ForegroundColor Yellow
try {
    $Apply = $rule.ApplyRule()
}
catch {
    write-warning $_
}
