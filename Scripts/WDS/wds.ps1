$Domain="$env:userdomain"
Install-WindowsFeature -name WDS -IncludeManagementTools
WDSUTIL /initialize-server /reminst:"D:\Reminst" /Authorize
Start-Service -displayname "Windows Deployment Services Server"
WDSUTIL /Set-Server /AnswerClients:all
WDSUTIL /Set-Server /PxepromptPolicy /Known:OptOut /New:OptOut
WDSUTIL /Set-Server /UseDHCPPorts:no /DHCPoption60:yes
WDSUTIL /Set-Server /NewMachineNamingPolicy:"$Domain"-%#
New-WdsInstallImageGroup -Name "Clients"
New-WdsInstallImageGroup -Name "Servers"
