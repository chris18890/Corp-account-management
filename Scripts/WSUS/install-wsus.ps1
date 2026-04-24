#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

# Mode: Standalone / No Shared Modules

[CmdletBinding()]
param([Parameter(Mandatory)][string]$Drive)
Add-Type -AssemblyName Microsoft.UpdateServices.Administration

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
#=========================================
if (!(TEST-PATH "$Drive\$RootShare")) {
    New-Item -Path $Drive -Name $RootShare -ItemType Directory -force
}
Install-WindowsFeature -Name UpdateServices -IncludeManagementTools
cd "C:\Program Files\Update Services\Tools\"
.\wsusutil.exe postinstall CONTENT_DIR=$Drive\$RootShare
if ($LASTEXITCODE -ne 0) { throw "wsusutil postinstall failed — exit code $LASTEXITCODE" }
Import-GPO -BackupGpoName $RootShare -TargetName $RootShare -path $GPOLocation -CreateIfNeeded
try {
    New-GPLink -name $RootShare -target $EndPath -LinkEnabled Yes -enforced yes
    Write-Host "Linked $RootShare to $EndPath"
} catch {
    $ex = $_.Exception
    if ($ex.Message -match "already exists") {
        Write-Host "'$RootShare' already linked to $EndPath" -ForegroundColor Green
    } else {
        throw
    }
}
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
    Write-Host "." -NoNewline
    Start-Sleep -Seconds 5
}
write-host " "
Write-Host "Sync is done." -ForegroundColor Green
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
