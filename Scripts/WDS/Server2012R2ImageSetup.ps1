Import-WdsBootImage -Path "E:\Sources\boot.wim" -NewImageName "Server 2012 R2 Boot x64"
Import-WdsInstallImage -Path "E:\Sources\install.wim" -ImageName "Windows Server 2012 R2 SERVERDATACENTER" -ImageGroup "Servers" -NewImageName "Server 2012 R2 Datacenter" -UnattendFile "C:\Scripts\WDS\Server2012R2Unattended.xml"
Import-WdsInstallImage -Path "E:\Sources\install.wim" -ImageName "Windows Server 2012 R2 SERVERDATACENTERCORE" -ImageGroup "Servers" -NewImageName "Server 2012 R2 Datacenter Core" -UnattendFile "C:\Scripts\WDS\Server2012R2Unattended.xml"
