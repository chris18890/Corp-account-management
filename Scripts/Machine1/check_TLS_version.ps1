Function Get-TlsRegValue {
    [CmdletBinding()]
    Param (
        # Registry Path
        [Parameter(Mandatory=$true,Position=0)]
        [string]
        $RegPath,
        # Registry Name
        [Parameter(Mandatory=$true,Position=1)]
        [string]
        $RegName
    )
    $regItem = Get-ItemProperty -Path $RegPath -Name $RegName -ErrorAction Ignore
    $output = "" | select Path,Name,Value
    $output.Path = $RegPath
    $output.Name = $RegName
    If ($regItem -eq $null) {
        $output.Value = "Not Found"
    } Else {
        $output.Value = $regItem.$RegName
    }
    $output
}

$regSettings = @()
$regKey = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727'
$regSettings += Get-TlsRegValue $regKey 'SystemDefaultTlsVersions'
$regSettings += Get-TlsRegValue $regKey 'SchUseStrongCrypto'

$regKey = 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727'
$regSettings += Get-TlsRegValue $regKey 'SystemDefaultTlsVersions'
$regSettings += Get-TlsRegValue $regKey 'SchUseStrongCrypto'

$regKey = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319'
$regSettings += Get-TlsRegValue $regKey 'SystemDefaultTlsVersions'
$regSettings += Get-TlsRegValue $regKey 'SchUseStrongCrypto'

$regKey = 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319'
$regSettings += Get-TlsRegValue $regKey 'SystemDefaultTlsVersions'
$regSettings += Get-TlsRegValue $regKey 'SchUseStrongCrypto'

$regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 2.0\Server'
$regSettings += Get-TlsRegValue $regKey 'Enabled'
$regSettings += Get-TlsRegValue $regKey 'DisabledByDefault'

$regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 2.0\Client'
$regSettings += Get-TlsRegValue $regKey 'Enabled'
$regSettings += Get-TlsRegValue $regKey 'DisabledByDefault'

$regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server'
$regSettings += Get-TlsRegValue $regKey 'Enabled'
$regSettings += Get-TlsRegValue $regKey 'DisabledByDefault'

$regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Client'
$regSettings += Get-TlsRegValue $regKey 'Enabled'
$regSettings += Get-TlsRegValue $regKey 'DisabledByDefault'

$regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server'
$regSettings += Get-TlsRegValue $regKey 'Enabled'
$regSettings += Get-TlsRegValue $regKey 'DisabledByDefault'

$regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client'
$regSettings += Get-TlsRegValue $regKey 'Enabled'
$regSettings += Get-TlsRegValue $regKey 'DisabledByDefault'

$regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server'
$regSettings += Get-TlsRegValue $regKey 'Enabled'
$regSettings += Get-TlsRegValue $regKey 'DisabledByDefault'

$regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client'
$regSettings += Get-TlsRegValue $regKey 'Enabled'
$regSettings += Get-TlsRegValue $regKey 'DisabledByDefault'

$regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server'
$regSettings += Get-TlsRegValue $regKey 'Enabled'
$regSettings += Get-TlsRegValue $regKey 'DisabledByDefault'

$regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client'
$regSettings += Get-TlsRegValue $regKey 'Enabled'
$regSettings += Get-TlsRegValue $regKey 'DisabledByDefault'

$regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Server'
$regSettings += Get-TlsRegValue $regKey 'Enabled'
$regSettings += Get-TlsRegValue $regKey 'DisabledByDefault'

$regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Client'
$regSettings += Get-TlsRegValue $regKey 'Enabled'
$regSettings += Get-TlsRegValue $regKey 'DisabledByDefault'

$regSettings
