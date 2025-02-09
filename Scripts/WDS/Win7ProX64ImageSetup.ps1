Import-WdsBootImage -Path "E:\Sources\boot.wim" -NewImageName "Windows 7 Boot x64"
Import-WdsInstallImage -Path "E:\Sources\install.wim" -ImageName "Windows 7 Professional" -ImageGroup "Clients" -NewImageName "Windows 7 Professional x64" -UnattendFile "C:\Scripts\WDS\Win7ProX64Unattended.xml"
