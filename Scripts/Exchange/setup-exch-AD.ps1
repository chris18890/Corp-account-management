#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

$Domain = "$env:userdomain"
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$Member = "$env:username"
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptTitle = "$Domain Exchange Prereq Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
$LogFileName = $Domain + "_exchange_prereq_log-$(Get-Date -Format 'yyyyMMdd')"
$LogIndex = 0
while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
    $LogIndex ++
}
$LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"

Write-Log ("=" * 80)
Write-Log "Log file is '$LogFile'"
Write-Log ("=" * 80)
Write-Log "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log "DC being used is '$DCHostName'"
Write-Log "Script path is '$ScriptPath'"
Write-Log "$ScriptTitle"
Write-Log "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log ("=" * 80)
Write-Log ""
$requiredGroups = @('Enterprise Admins', 'Schema Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

Write-Log "Intalling IIS URL Rewrite"
msiexec.exe /i "\\$DNSSuffix\Store\Software\rewrite_amd64_en-US.msi" /quiet
if ($LASTEXITCODE -ne 0) { throw "msiexec (URL Rewrite) failed — exit code $LASTEXITCODE" }
Write-Log "Intalling Visual C++ Runtime"
Start-Process -Filepath "\\$DNSSuffix\Store\Software\vcredist_x64_2013.exe" -Argumentlist "/Q" -wait
if ($LASTEXITCODE -ne 0) { throw "vcredist install failed — exit code $LASTEXITCODE" }
Write-Log "Intalling Unified Communications Managed API"
.\UCMARedist\Setup.exe -q
if ($LASTEXITCODE -ne 0) { throw "UCMA install failed — exit code $LASTEXITCODE" }
Write-Log "Launching Exchange Setup with the /PrepareAD switch"
try {
    .\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAD /OrganizationName:"$Domain"
    if ($LASTEXITCODE -ne 0) { throw "Exchange PrepareAD failed — exit code $LASTEXITCODE. Enterprise/Schema Admins NOT removed." }
} finally {
    try {
        Remove-ADGroupMember -Identity "Enterprise Admins" -Members $Member -Confirm:$False -Server $DCHostName
        Write-Log "Removed $Member from Enterprise Admins"
    } catch {
        $ex = $_.Exception
        Write-Log "ERROR: $($ex.Message)" -ForegroundColor Red
    }
    try {
        Remove-ADGroupMember -Identity "Schema Admins" -Members $Member -Confirm:$False -Server $DCHostName
        Write-Log "Removed $Member from Schema Admins"
    } catch {
        $ex = $_.Exception
        Write-Log "ERROR: $($ex.Message)" -ForegroundColor Red
    }
}
Restart-Computer
