# Run as Administrator. Lets Power BI Desktop load SynologyOffice.pqx while the data
# extension security level stays at "Recommended": adds the signing certificate's
# thumbprint to the TrustedCertificateThumbprints policy (other thumbprints are kept).
#
# With no -Thumbprint it auto-detects the connector's code-signing certificate
# ("CN=SynologyOffice Power BI Connector*") in the current user's store — the same one
# build.ps1 -Sign uses. Pass -Thumbprint to trust a specific certificate instead.
param([string]$Thumbprint)

$ErrorActionPreference = 'Stop'
$identity = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $identity.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run this script in PowerShell started "As Administrator".'
}

if (-not $Thumbprint) {
    $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
        Where-Object { $_.Subject -like 'CN=SynologyOffice Power BI Connector*' -and $_.NotAfter -gt (Get-Date) } |
        Sort-Object NotAfter -Descending | Select-Object -First 1
    if (-not $cert) {
        throw 'Signing certificate "CN=SynologyOffice Power BI Connector" not found in Cert:\CurrentUser\My. Create it with: build.ps1 -NewCert (or pass -Thumbprint).'
    }
    $Thumbprint = $cert.Thumbprint
}

$key = 'HKLM:\SOFTWARE\Policies\Microsoft\Power BI Desktop'
New-Item -Path $key -Force | Out-Null
$current = @((Get-ItemProperty -Path $key -Name TrustedCertificateThumbprints -ErrorAction SilentlyContinue).TrustedCertificateThumbprints)
$values = @($current | Where-Object { $_ }) + $Thumbprint.ToUpper() | Select-Object -Unique
New-ItemProperty -Path $key -Name TrustedCertificateThumbprints -PropertyType MultiString -Value $values -Force | Out-Null
Write-Host "TrustedCertificateThumbprints = $($values -join ', ')"
Write-Host 'Restart Power BI Desktop.'
