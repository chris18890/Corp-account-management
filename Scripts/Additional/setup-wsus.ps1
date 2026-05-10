#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Drive
    ,[Parameter(Mandatory)][string]$Domain
    ,[string]$DomainSuffix
    ,[string]$LogFile
)

Set-StrictMode -Version Latest
Add-Type -AssemblyName Microsoft.UpdateServices.Administration

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

$Drive = $Drive.TrimEnd(':') + ':'
$ShareName = $Env.Shares.WSUS
$FeatureName = @("UpdateServices")
$ScriptPath = $PSScriptRoot
$ScriptTitle = "$Domain network setup Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_WSUS_setup_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-LogFile -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-LogFile -LogFile $LogFile -LogString "$ScriptTitle"
Write-LogFile -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Note - Script must be run TWICE on this server - first to do the domain join then again to set up WSUS"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "

Foreach ($Feature in $FeatureName){
    if (Get-WindowsFeature -Name $Feature | Where-Object InstallState -Eq Installed) {
        Write-LogFile -LogFile $LogFile -LogString "$Feature is already installed" -ForegroundColor Green
    } else {
        Write-LogFile -LogFile $LogFile -LogString "installing $Feature"
        install-windowsfeature -IncludeManagementTools -Name $Feature
        Write-LogFile -LogFile $LogFile -LogString "installed $Feature"
    }
}
if ((Get-CimInstance win32_computersystem).partofdomain -eq $false) {
    If (!$DomainSuffix) {
        $DomainSuffix = READ-HOST 'Enter a public FQDN- '
    }
    $DNSSuffix = "$Domain.$DomainSuffix"
    Add-Computer -DomainName "$DNSSuffix" -Restart
} else {
    if ((Get-CimInstance win32_computersystem).partofdomain -eq $true) {
        $EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
        $DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
        $DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
        $GPOLocation = Join-Path (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "Prelim") "GPOs"
        if (!(TEST-PATH "$Drive\$ShareName")) {
            New-Item -Path $Drive -Name $ShareName -ItemType Directory -force
        } else {
            Write-LogFile -LogFile $LogFile -LogString "$Drive\$ShareName already exists" -ForegroundColor Green
        }
        & "C:\Program Files\Update Services\Tools\wsusutil.exe" postinstall CONTENT_DIR=$Drive\$ShareName
        if ($LASTEXITCODE -ne 0) { throw "wsusutil postinstall failed - exit code $LASTEXITCODE" }
        Import-GPO -BackupGpoName $ShareName -TargetName $ShareName -path $GPOLocation -CreateIfNeeded -Server $DCHostName
        Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName $ShareName -GPOTarget $EndPath
        try {
            Write-LogFile -LogFile $LogFile -LogString "Updating permissions on $ShareName"
            Set-GPPermission -Name $ShareName -PermissionLevel GpoEditDeleteModifySecurity -TargetName "$($Env.Groups.TaskPrefix)AD_GPO_Admins" -TargetType Group -Server $DCHostName
            Write-LogFile -LogFile $LogFile -LogString "Updated permissions on $ShareName"
        } catch {
            throw
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
        Write-LogFile -LogFile $LogFile -LogString "Sync is done." -ForegroundColor Green
        # Configure the Platforms that we want WSUS to receive updates
        Write-LogFile -LogFile $LogFile -LogString "Setting WSUS Products"
        Get-WsusProduct | where-Object {
            $_.Product.Title -in ($Env.WSUS.Products)
        } | Set-WsusProduct
        # Configure the Classifications
        Write-LogFile -LogFile $LogFile -LogString "Setting WSUS Classifications"
        Get-WsusClassification | Where-Object {
            $_.Classification.Title -in ($Env.WSUS.Classifications)
        } | Set-WsusClassification
        # Configure Synchronizations
        Write-LogFile -LogFile $LogFile -LogString "Enabling WSUS Automatic Synchronisation"
        $subscription.SynchronizeAutomatically=$true
        # Set synchronization scheduled for midnight each night
        $subscription.SynchronizeAutomaticallyTimeOfDay= (New-TimeSpan -Hours 0)
        $subscription.NumberOfSynchronizationsPerDay=1
        $subscription.Save()
        # Kick off a synchronization
        $subscription.StartSynchronization()
        # Configure Default Approval Rule
        Write-LogFile -LogFile $LogFile -LogString "Configuring default automatic approval rule"
        $rule = $wsus.GetInstallApprovalRules() | Where-Object {
            $_.Name -eq "Default Automatic Approval Rule"
        }
        $class = $wsus.GetUpdateClassifications() | Where-Object {
            $_.Title -In ($Env.WSUS.Classifications)
        }
        $class_coll = New-Object Microsoft.UpdateServices.Administration.UpdateClassificationCollection
        $class_coll.AddRange($class)
        $rule.SetUpdateClassifications($class_coll)
        $rule.Enabled = $True
        $rule.Save()
        
        # Run Default Approval Rule
        Write-LogFile -LogFile $LogFile -LogString "Running Default Approval Rule"
        Write-LogFile -LogFile $LogFile -LogString " >This step may timeout, but the rule will be applied and the script will continue" -ForegroundColor Yellow
        try {
            $rule.ApplyRule()
        }
        catch {
            Write-LogFile -LogFile $LogFile -LogString $_
        }
    }
}
