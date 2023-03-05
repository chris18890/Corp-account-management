param([string]$Domain)
If (!$Domain) {
    $Domain = READ-HOST 'Enter a NETBIOS Domain name- '
}
$ServerName = "$env:computername"
Install-WindowsFeature -IncludeManagementTools Adcs-Cert-Authority, rsat-ad-powershell
$FQDN = [System.Net.Dns]::GetHostByName(($ServerNameName)).Hostname
$DNSSuffix = $FQDN.Substring($FQDN.IndexOf(".") + 1)
$EndPath = "dc=" + ($DNSSuffix.Replace(".", ",dc="))
$httpCRLPath = "ca.$DNSSuffix"
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
CRLPeriodUnits=10
CRLDeltaPeriod=Days
CRLDeltaPeriodUnits=0
LoadDefaultTemplates=0
AlternateSignatureAlgorithm=1
"@
$CAPolicyInf | Out-File "C:\Windows\CAPolicy.inf" -Encoding utf8 -Force | Out-Null

Install-AdcsCertificationAuthority -CAType StandaloneRootCa -CryptoProviderName "RSA#Microsoft Software Key Storage Provider" -KeyLength 4096 -HashAlgorithmName SHA256 -ValidityPeriod Years -ValidityPeriodUnits 20 -CACommonName $ServerName -CADistinguishedNameSuffix $EndPath
Get-CACrlDistributionPoint | Remove-CACrlDistributionPoint -Force | Out-Null
Add-CACRLDistributionPoint -Uri "$env:windir\system32\CertSrv\CertEnroll\$ServerName.crl" -PublishToServer -PublishDeltaToServer -Force | Out-Null
Add-CACRLDistributionPoint -Uri "C:\CAConfig\$ServerName.crl" -PublishToServer -PublishDeltaToServer -Force | Out-Null
Add-CACRLDistributionPoint -Uri "http://$httpCRLPath/certenroll/$ServerName.crl" -AddToCertificateCDP -AddToFreshestCrl -Force | Out-Null
Get-CAAuthorityInformationAccess | where { $_.Uri -like '*ldap*' -or $_.Uri -like '*http*' -or $_.Uri -like '*file*' } | Remove-CAAuthorityInformationAccess -Force | Out-Null
Add-CAAuthorityInformationAccess -Uri "http://$httpCRLPath/certenroll/$ServerName.$DNSSuffix_$RootCAName.crt" -AddToCertificateAia -Force  | Out-Null
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
Start-Sleep 5
certutil -crl | Out-Null

# certreq -submit "\\tsclient\c\CA\$InterCAName.$DNSSuffix._$Domain-$InterCAName-CA.req"
# certreq -retrieve 2 "\\tsclient\c\CA\$InterCAName.$DNSSuffix._$Domain-$InterCAName-CA.crt"
