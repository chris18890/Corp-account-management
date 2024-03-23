Import-WdsBootImage -Path "E:\Sources\boot.wim" -NewImageName "Server 2022 Boot x64"
Import-WdsInstallImage -Path "E:\Sources\install.wim" -ImageName "Windows Server 2022 SERVERDATACENTER" -ImageGroup "Servers" -NewImageName "Server 2022 Datacenter" -UnattendFile "C:\Scripts\WDS\Server2022Unattended.xml"
Import-WdsInstallImage -Path "E:\Sources\install.wim" -ImageName "Windows Server 2022 SERVERDATACENTERCORE" -ImageGroup "Servers" -NewImageName "Server 2022 Datacenter Core" -UnattendFile "C:\Scripts\WDS\Server2022Unattended.xml"
