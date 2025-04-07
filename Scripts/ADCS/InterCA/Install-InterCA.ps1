$Domain = "$env:userdomain"
$ServerName = "$env:computername"
Install-WindowsFeature -IncludeManagementTools Adcs-Cert-Authority, ADCS-Web-Enrollment, rsat-ad-powershell, rsat-dns-server
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$DNSServer = (Get-ADDomainController).HostName
$httpCRLPath = "ca.$DNSSuffix"
$Drive = "C:"
$ShareName = "CertEnroll"
$ManagementServer = "$Domain-RTR"
$RootCAName = "$Domain-RootCA"

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
$CAPolicyInf | Out-File "C:\Windows\CAPolicy.inf" -Encoding utf8 -Force | Out-Null

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
New-SmbShare -Name $ShareName -Path "$Drive\$ShareName" -FullAccess "authenticated users"
New-WebVirtualDirectory -Site "Default Web Site" -Name "$ShareName" -PhysicalPath "$Drive\$ShareName"
New-WebBinding -Name "Default Web Site" -IPAddress "*" -Port 80 -HostHeader "$httpCRLPath"

cd inetsrv\
.\Appcmd set config "Default Web Site" /section:system.webServer/Security/requestFiltering -allowDoubleEscaping:True
IISReset
cd

Add-DnsServerResourceRecordCName -ComputerName $DNSServer -Name "ca" -HostNameAlias "$ServerName.$DNSSuffix" -ZoneName "$DNSSuffix"

pause

robocopy -e -s "\\$ManagementServer\c$\ca" "\\$ServerName\CertEnroll"
certutil -f -dspublish "\\$ServerName\CertEnroll\$RootCAName.crl" $RootCAName
certutil -f -dspublish "\\$ServerName\CertEnroll\$RootCAName.$DNSSuffix_$RootCAName.crt" RootCA
certutil -addstore -f root "\\$ServerName\CertEnroll\$RootCAName.crl"
certutil -addstore -f root "\\$ServerName\CertEnroll\$RootCAName.$DNSSuffix_$RootCAName.crt" 
Install-AdcsCertificationAuthority -CAType EnterpriseSubordinateCa -CryptoProviderName "RSA#Microsoft Software Key Storage Provider" -KeyLength 4096 -HashAlgorithmName SHA256 -CACommonName $ServerName -CADistinguishedNameSuffix $EndPath -OutputCertRequestFile "$Drive\$ShareName\$ServerName.$DNSSuffix._$Domain-$ServerName-CA.req"
Install-AdcsWebEnrollment

pause

Restart-Service certsvc | Out-Null

Get-CACrlDistributionPoint | Remove-CACrlDistributionPoint -Force | Out-Null
Add-CACRLDistributionPoint -Uri "$env:windir\system32\CertSrv\CertEnroll\$RootCAName.crl" -PublishToServer -PublishDeltaToServer -Force | Out-Null
Add-CACRLDistributionPoint -Uri "C:\CAConfig\$RootCAName.crl" -PublishToServer -PublishDeltaToServer -Force | Out-Null
Add-CACRLDistributionPoint -Uri "http://$httpCRLPath/certenroll/$RootCAName.crl" -AddToCertificateCDP -AddToFreshestCrl -Force | Out-Null
Get-CAAuthorityInformationAccess | where { $_.Uri -like '*ldap*' -or $_.Uri -like '*http*' -or $_.Uri -like '*file*' } | Remove-CAAuthorityInformationAccess -Force | Out-Null
Add-CAAuthorityInformationAccess -Uri "http://$httpCRLPath/certenroll/$RootCAName.$DNSSuffix_$RootCAName.crt" -AddToCertificateAia -Force | Out-Null
certutil.exe -setreg DSDomainDN "$EndPath" | Out-Null
certutil.exe -setreg CADSConfigDN "CN=Configuration,$EndPath" | Out-Null
certutil.exe -setreg CA\CRLPeriodUnits 10 | Out-Null
certutil.exe -setreg CA\CRLPeriod "Years" | Out-Null
certutil.exe -setreg CA\ValidityPeriodUnits 10 | Out-Null
certutil.exe -setreg CA\ValidityPeriod "Years" | Out-Null
certutil.exe -setreg CA\CRLOverlapPeriodUnits 3 | Out-Null
certutil.exe -setreg CA\CRLOverlapPeriod "Weeks" | Out-Null
certutil.exe -setreg CA\AuditFilter 127 | Out-Null
Restart-Service certsvc | Out-Null
certutil -crl | Out-Null
