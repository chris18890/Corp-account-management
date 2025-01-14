Import-WdsBootImage -Path "E:\Sources\boot.wim" -NewImageName "Server 2016 Boot x64"
Import-WdsInstallImage -Path "E:\Sources\install.wim" -ImageName "Windows Server 2016 SERVERDATACENTER" -ImageGroup "Servers" -NewImageName "Server 2016 Datacenter" -UnattendFile "C:\Scripts\WDS\Server2016Unattended.xml"
Import-WdsInstallImage -Path "E:\Sources\install.wim" -ImageName "Windows Server 2016 SERVERDATACENTERCORE" -ImageGroup "Servers" -NewImageName "Server 2016 Datacenter Core" -UnattendFile "C:\Scripts\WDS\Server2016Unattended.xml"
