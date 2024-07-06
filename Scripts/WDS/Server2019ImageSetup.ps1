Import-WdsBootImage -Path "E:\Sources\boot.wim" -NewImageName "Server 2019 Boot x64"
Import-WdsInstallImage -Path "E:\Sources\install.wim" -ImageName "Windows Server 2019 SERVERDATACENTER" -ImageGroup "Servers" -NewImageName "Server 2019 Datacenter" -UnattendFile "C:\Scripts\WDS\Server2019Unattended.xml"
Import-WdsInstallImage -Path "E:\Sources\install.wim" -ImageName "Windows Server 2019 SERVERDATACENTERCORE" -ImageGroup "Servers" -NewImageName "Server 2019 Datacenter Core" -UnattendFile "C:\Scripts\WDS\Server2019Unattended.xml"
