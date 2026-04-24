#Requires -RunAsAdministrator
#Requires -Modules ActiveDirectory, WebAdministration
Set-StrictMode -Version Latest

# Execution Tier: Tier-0
# Mode: Standalone / No Shared Modules

$Domain = "$env:userdomain"
$ServerName = "$env:computername"
Install-WindowsFeature -IncludeManagementTools Adcs-Cert-Authority, ADCS-Web-Enrollment, rsat-ad-powershell, rsat-dns-server
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$httpCRLPath = "ca.$DNSSuffix"
$Drive = "C:"
$ShareName = "CertEnroll"
$ManagementServer = "$Domain-RTR"
$RootCAName = "$Domain-RootCA"
$requiredGroups = @('Enterprise Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Host "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

$CAPolicyInf = @"
[Version]
Signature="`$Windows NT$"
[PolicyStatementExtension]
Policies=InternalPolicy
[InternalPolicy]
OID= 1.3.6.1.4.1.2.5.29.32.0
Notice="Legal Policy Statement"
URL=http://$httpCRLPath/pki/cps.html
[Certsrv_Server]
RenewalKeyLength=4096
RenewalValidityPeriod=Years
RenewalValidityPeriodUnits=5
LoadDefaultTemplates=1
AlternateSignatureAlgorithm=1
"@
$CAPolicyInf | Out-File "C:\Windows\CAPolicy.inf" -Encoding utf8 -Force

New-Item "$Drive\$ShareName" -type directory -force
$Acl = Get-Acl "$Drive\$ShareName"
$isProtected = $true
$preserveInheritance = $false
$Acl.SetAccessRuleProtection($isProtected, $preserveInheritance)
$Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Cert Publishers","Modify","ContainerInherit, ObjectInherit", "None", "Allow")
$Acl.SetAccessRule($Ar)
$Ar = New-Object system.security.accesscontrol.filesystemaccessrule("ADM_Task_DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
$Acl.SetAccessRule($Ar)
$Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
$Acl.SetAccessRule($Ar)
$Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
$Acl.SetAccessRule($Ar)
Set-Acl "$Drive\$ShareName" $Acl
if (-not (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
    New-SmbShare -Name $ShareName -Path "$Drive\$ShareName" -FullAccess "Administrators", "SYSTEM" -ChangeAccess "authenticated users"
} else {
    Write-Host "SMB share '$ShareName' already exists" -ForegroundColor Green
}
if (-not (Get-WebVirtualDirectory -Site "Default Web Site" -Name $ShareName -ErrorAction SilentlyContinue)) {
    New-WebVirtualDirectory -Site "Default Web Site" -Name "$ShareName" -PhysicalPath "$Drive\$ShareName"
} else {
    Write-Host "Virtual directory '$ShareName' already exists" -ForegroundColor Green
}
if (-not (Get-WebBinding -Name "Default Web Site" -HostHeader $httpCRLPath -ErrorAction SilentlyContinue)) {
    New-WebBinding -Name "Default Web Site" -IPAddress "*" -Port 80 -HostHeader "$httpCRLPath"
} else {
    Write-Host "Web binding for '$httpCRLPath' already exists" -ForegroundColor Green
}

Set-Location "$env:SystemRoot\system32\inetsrv"
.\Appcmd set config "Default Web Site/CertEnroll" /section:system.webServer/Security/requestFiltering -allowDoubleEscaping:True
IISReset
cd

if (-not (Get-DnsServerResourceRecord -ComputerName $DCHostName -ZoneName $DNSSuffix -Name "ca" -ErrorAction SilentlyContinue)) {
    Add-DnsServerResourceRecordCName -ComputerName $DCHostName -Name "ca" -HostNameAlias "$ServerName.$DNSSuffix" -ZoneName "$DNSSuffix"
} else {
    Write-Host "DNS CNAME 'ca' already exists" -ForegroundColor Green
}

Install-AdcsCertificationAuthority -CAType EnterpriseSubordinateCa -CryptoProviderName "RSA#Microsoft Software Key Storage Provider" -KeyLength 4096 -HashAlgorithmName SHA256 -CACommonName $ServerName -CADistinguishedNameSuffix $EndPath -OutputCertRequestFile "$Drive\$ShareName\$ServerName.$DNSSuffix._$Domain-$ServerName-CA.req"
Install-AdcsWebEnrollment

Read-Host -Prompt "Press Enter after copying the CSR to the Root CA"

robocopy "\\$ManagementServer\c$\ca" "\\$ServerName\CertEnroll" /e
if ($LASTEXITCODE -ge 8) { throw "robocopy failed — exit code $LASTEXITCODE" }
certutil -f -dspublish "\\$ServerName\CertEnroll\$RootCAName.crl" $RootCAName
if ($LASTEXITCODE -ne 0) { throw "certutil -dspublish (CRL) failed — exit code $LASTEXITCODE" }
certutil -f -dspublish "\\$ServerName\CertEnroll\$RootCAName.$($DNSSuffix)_$RootCAName.crt" RootCA
if ($LASTEXITCODE -ne 0) { throw "certutil -dspublish (RootCA CRT) failed — exit code $LASTEXITCODE" }
certutil -addstore -f root "\\$ServerName\CertEnroll\$RootCAName.crl"
if ($LASTEXITCODE -ne 0) { throw "certutil -addstore (CRL) failed — exit code $LASTEXITCODE" }
certutil -addstore -f root "\\$ServerName\CertEnroll\$RootCAName.$($DNSSuffix)_$RootCAName.crt" 
if ($LASTEXITCODE -ne 0) { throw "certutil -addstore (CRT) failed — exit code $LASTEXITCODE" }

Read-Host -Prompt "Press Enter after the CSR has been signed by the Root CA and the .crt copied back to \\$ServerName\CertEnroll"

Restart-Service certsvc

Get-CACrlDistributionPoint | Remove-CACrlDistributionPoint -Force
Add-CACRLDistributionPoint -Uri "$env:windir\system32\CertSrv\CertEnroll\$RootCAName.crl" -PublishToServer -PublishDeltaToServer -Force
Add-CACRLDistributionPoint -Uri "C:\CAConfig\$RootCAName.crl" -PublishToServer -PublishDeltaToServer -Force
Add-CACRLDistributionPoint -Uri "http://$httpCRLPath/certenroll/$RootCAName.crl" -AddToCertificateCDP -AddToFreshestCrl -Force
Get-CAAuthorityInformationAccess | where { $_.Uri -like '*ldap*' -or $_.Uri -like '*http*' -or $_.Uri -like '*file*' } | Remove-CAAuthorityInformationAccess -Force
Add-CAAuthorityInformationAccess -Uri "http://$httpCRLPath/certenroll/$RootCAName.$($DNSSuffix)_$RootCAName.crt" -AddToCertificateAia -Force
certutil.exe -setreg DSDomainDN "$EndPath"
if ($LASTEXITCODE -ne 0) { throw "certutil -setreg DSDomainDN failed — exit code $LASTEXITCODE" }
certutil.exe -setreg CADSConfigDN "CN=Configuration,$EndPath"
if ($LASTEXITCODE -ne 0) { throw "certutil -setreg CADSConfigDN failed — exit code $LASTEXITCODE" }
certutil.exe -setreg CA\CRLPeriodUnits 1
if ($LASTEXITCODE -ne 0) { throw "certutil -setreg CRLPeriodUnits failed — exit code $LASTEXITCODE" }
certutil.exe -setreg CA\CRLPeriod "Years"
if ($LASTEXITCODE -ne 0) { throw "certutil -setreg CRLPeriod failed — exit code $LASTEXITCODE" }
certutil.exe -setreg CA\ValidityPeriodUnits 5
if ($LASTEXITCODE -ne 0) { throw "certutil -setreg ValidityPeriodUnits failed — exit code $LASTEXITCODE" }
certutil.exe -setreg CA\ValidityPeriod "Years"
if ($LASTEXITCODE -ne 0) { throw "certutil -setreg ValidityPeriod failed — exit code $LASTEXITCODE" }
certutil.exe -setreg CA\CRLOverlapPeriodUnits 3
if ($LASTEXITCODE -ne 0) { throw "certutil -setreg CRLOverlapPeriodUnits failed — exit code $LASTEXITCODE" }
certutil.exe -setreg CA\CRLOverlapPeriod "Weeks"
if ($LASTEXITCODE -ne 0) { throw "certutil -setreg CRLOverlapPeriod failed — exit code $LASTEXITCODE" }
certutil.exe -setreg CA\AuditFilter 127
if ($LASTEXITCODE -ne 0) { throw "certutil -setreg AuditFilter failed — exit code $LASTEXITCODE" }
Restart-Service certsvc
certutil -crl
if ($LASTEXITCODE -ne 0) { throw "certutil -crl failed — exit code $LASTEXITCODE" }
