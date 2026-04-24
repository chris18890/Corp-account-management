#Requires -RunAsAdministrator

[CmdletBinding()]
param([Parameter(Mandatory)][string]$Drive)

Set-StrictMode -Version Latest

# Mode: Standalone / No Shared Modules

$Domain = "$env:userdomain"
$Drive = $Drive.TrimEnd(':') + ':'
$RootShare = "Reminst"

Install-WindowsFeature -name WDS -IncludeManagementTools
WDSUTIL /initialize-server /reminst:"$Drive\$RootShare" /Authorize
if ($LASTEXITCODE -ne 0) { throw "WDSUTIL command to configure server failed - exit code $LASTEXITCODE" }
Start-Service -displayname "Windows Deployment Services Server"
WDSUTIL /Set-Server /AnswerClients:all
if ($LASTEXITCODE -ne 0) { throw "WDSUTIL command to set client answer policy failed - exit code $LASTEXITCODE" }
WDSUTIL /Set-Server /PxepromptPolicy /Known:OptOut /New:OptOut
if ($LASTEXITCODE -ne 0) { throw "WDSUTIL command to set PXE prompt policy failed - exit code $LASTEXITCODE" }
WDSUTIL /Set-Server /UseDHCPPorts:no /DHCPoption60:yes
if ($LASTEXITCODE -ne 0) { throw "WDSUTIL command to set DHCP options failed - exit code $LASTEXITCODE" }
WDSUTIL /Set-Server /NewMachineNamingPolicy:"$Domain"-%#
if ($LASTEXITCODE -ne 0) { throw "WDSUTIL command to set naming policy failed - exit code $LASTEXITCODE" }
New-WdsInstallImageGroup -Name "Clients"
New-WdsInstallImageGroup -Name "Servers"
