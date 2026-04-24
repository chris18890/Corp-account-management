#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

# Execution Tier: Tier-0
# Mode: Standalone / No Shared Modules

param([string]$Domain,[string]$DomainSuffix)

If (!$Domain) {
    $Domain = READ-HOST 'Enter a NETBIOS Domain name- '
}
If (!$DomainSuffix) {
    $DomainSuffix = READ-HOST 'Enter a public FQDN- '
}

$ServerName = "$env:computername"
Install-WindowsFeature -IncludeManagementTools Adcs-Cert-Authority, rsat-ad-powershell
$DNSName = "$Domain.$DomainSuffix"
if ([string]::IsNullOrWhiteSpace($DomainSuffix)) {
    throw "DomainSuffix parameter is required (e.g. 'company.com')"
}
$EndPath = "dc=" + ($DNSName.Replace(".", ",dc="))
$httpCRLPath = "ca.$DNSName"
$ManagementServer = "$Domain-RTR"
$RootCAName = "$Domain-RootCA"
$InterCAName = "$Domain-InterCA"

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
RenewalValidityPeriodUnits=10
CRLPeriod=Years
CRLPeriodUnits=1
CRLDeltaPeriod=Days
CRLDeltaPeriodUnits=0
LoadDefaultTemplates=0
AlternateSignatureAlgorithm=1
"@
$CAPolicyInf | Out-File "C:\Windows\CAPolicy.inf" -Encoding utf8 -Force

Install-AdcsCertificationAuthority -CAType StandaloneRootCa -CryptoProviderName "RSA#Microsoft Software Key Storage Provider" -KeyLength 4096 -HashAlgorithmName SHA256 -ValidityPeriod Years -ValidityPeriodUnits 20 -CACommonName $ServerName -CADistinguishedNameSuffix $EndPath
Get-CACrlDistributionPoint | Remove-CACrlDistributionPoint -Force
Add-CACRLDistributionPoint -Uri "$env:windir\system32\CertSrv\CertEnroll\$ServerName.crl" -PublishToServer -PublishDeltaToServer -Force
Add-CACRLDistributionPoint -Uri "C:\CAConfig\$ServerName.crl" -PublishToServer -PublishDeltaToServer -Force
Add-CACRLDistributionPoint -Uri "http://$httpCRLPath/certenroll/$ServerName.crl" -AddToCertificateCDP -AddToFreshestCrl -Force
Get-CAAuthorityInformationAccess | where { $_.Uri -like '*ldap*' -or $_.Uri -like '*http*' -or $_.Uri -like '*file*' } | Remove-CAAuthorityInformationAccess -Force
Add-CAAuthorityInformationAccess -Uri "http://$httpCRLPath/certenroll/$RootCAName.$($DNSName)_$RootCAName.crt" -AddToCertificateAia -Force 
certutil.exe -setreg DSDomainDN "$EndPath"
if ($LASTEXITCODE -ne 0) { throw "certutil -setreg DSDomainDN failed — exit code $LASTEXITCODE" }
certutil.exe -setreg CADSConfigDN "CN=Configuration,$EndPath"
if ($LASTEXITCODE -ne 0) { throw "certutil -setreg CADSConfigDN failed — exit code $LASTEXITCODE" }
certutil.exe -setreg CA\CRLPeriodUnits 1
if ($LASTEXITCODE -ne 0) { throw "certutil -setreg CRLPeriodUnits failed — exit code $LASTEXITCODE" }
certutil.exe -setreg CA\CRLPeriod "Years"
if ($LASTEXITCODE -ne 0) { throw "certutil -setreg CRLPeriod failed — exit code $LASTEXITCODE" }
certutil.exe -setreg CA\ValidityPeriodUnits 10
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
Start-Sleep 5
certutil -crl
if ($LASTEXITCODE -ne 0) { throw "certutil -crl failed — exit code $LASTEXITCODE" }

# Below commands are to sign the CSR from the Intermediate CA

# certreq -submit "\\tsclient\c\CA\$InterCAName.$DNSName._$Domain-$InterCAName-CA.req"
# certreq -retrieve 2 "\\tsclient\c\CA\$InterCAName.$DNSName._$Domain-$InterCAName-CA.crt"
