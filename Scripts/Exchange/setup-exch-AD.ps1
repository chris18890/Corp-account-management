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
$requiredGroups = @('Enterprise Admins', 'Schema Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

Write-Log -LogFile $LogFile -LogString "Intalling IIS URL Rewrite"
Start-Process msiexec /i "\\$DNSSuffix\Store\Software\rewrite_amd64_en-US.msi" -Argumentlist "/i /Quiet" -wait
Write-Log -LogFile $LogFile -LogString "Intalling Visual C++ Runtime"
Start-Process -Filepath "\\$DNSSuffix\Store\Software\vcredist_x64_2013.exe" -Argumentlist "/Q" -wait
Write-Log -LogFile $LogFile -LogString "Intalling Unified Communications Managed API"
Start-Process -Filepath ".\UCMARedist\Setup.exe" -Argumentlist "/Q" -wait
Write-Log -LogFile $LogFile -LogString "Launching Exchange Setup with the /PrepareAD switch"
try {
    .\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAD /OrganizationName:"$Domain"
    if ($LASTEXITCODE -ne 0) { throw "Exchange PrepareAD failed - exit code $LASTEXITCODE. Enterprise/Schema Admins NOT removed." }
} finally {
    try {
        Remove-ADGroupMember -Identity "Enterprise Admins" -Members $Member -Confirm:$False -Server $DCHostName
        Write-Log -LogFile $LogFile -LogString "Removed $Member from Enterprise Admins"
    } catch {
        $ex = $_.Exception
        Write-Log -LogFile $LogFile -LogString "ERROR: $($ex.Message)" -ForegroundColor Red
    }
    try {
        Remove-ADGroupMember -Identity "Schema Admins" -Members $Member -Confirm:$False -Server $DCHostName
        Write-Log -LogFile $LogFile -LogString "Removed $Member from Schema Admins"
    } catch {
        $ex = $_.Exception
        Write-Log -LogFile $LogFile -LogString "ERROR: $($ex.Message)" -ForegroundColor Red
    }
}
Restart-Computer
