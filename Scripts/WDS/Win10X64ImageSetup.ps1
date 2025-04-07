Import-WdsBootImage -Path "E:\Sources\boot.wim" -NewImageName "Windows 10 Boot x64"
Import-WdsInstallImage -Path "E:\Sources\install.wim" -ImageName "Windows 10 Pro" -ImageGroup "Clients" -NewImageName "Windows 10 Pro x64" -UnattendFile "C:\Scripts\WDS\Win10X64Unattended.xml"
