# Builds SynologyOffice.mez (a zip of the .pq, resources and icons).
#   -NewCert  creates a self-signed code-signing certificate
#             "CN=SynologyOffice Power BI Connector" in Cert:\CurrentUser\My if none
#             exists yet (valid 5 years). Run once before the first -Sign.
#   -Sign     also packs a signed SynologyOffice.pqx using that code-signing certificate.
#   -Install  copies the result (.pqx when signed, otherwise .mez) to the Power BI
#             Desktop custom connectors folder and removes the other variant there.
param(
    [switch]$Install,
    [switch]$Sign,
    [switch]$NewCert,
    [string]$SdkTools = (Join-Path $env:LOCALAPPDATA 'PowerQuerySdkTools\2.155.2')
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$out = Join-Path $root 'bin'
New-Item -ItemType Directory -Force $out | Out-Null

Add-Type -AssemblyName System.Drawing
foreach ($size in 16, 20, 24, 32, 40, 48, 64) {
    $png = Join-Path $root "SynologyOffice$size.png"
    if (Test-Path $png) { continue }
    $bmp = New-Object System.Drawing.Bitmap $size, $size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $green = [System.Drawing.Color]::FromArgb(255, 32, 158, 92)
    $g.FillRectangle((New-Object System.Drawing.SolidBrush $green), 0, 0, $size, $size)
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), ([Math]::Max(1, $size / 16))
    $step = $size / 4
    for ($i = 1; $i -lt 4; $i++) {
        $g.DrawLine($pen, $step * $i, $step, $step * $i, $size - $step / 2)
        $g.DrawLine($pen, $step / 2, $step * $i, $size - $step / 2, $step * $i)
    }
    $g.Dispose()
    $bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

$mez = Join-Path $out 'SynologyOffice.mez'
if (Test-Path $mez) { [IO.File]::Delete($mez) }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open($mez, 'Create')
try {
    $files = @('SynologyOffice.pq') +
        (Get-ChildItem $root -Filter 'resources*.resx' | ForEach-Object Name) +
        (Get-ChildItem $root -Filter 'SynologyOffice*.png' | ForEach-Object Name)
    foreach ($f in $files) {
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, (Join-Path $root $f), $f) | Out-Null
    }
} finally { $zip.Dispose() }
Write-Host "Built $mez"
$artifact = $mez

$FindCert = {
    Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
        Where-Object { $_.Subject -like 'CN=SynologyOffice Power BI Connector*' -and $_.NotAfter -gt (Get-Date) } |
        Sort-Object NotAfter -Descending | Select-Object -First 1
}

if ($NewCert) {
    $existing = & $FindCert
    if ($existing) {
        Write-Host "Signing certificate already exists: $($existing.Thumbprint)"
    } else {
        $created = New-SelfSignedCertificate -Type CodeSigningCert `
            -Subject 'CN=SynologyOffice Power BI Connector' `
            -KeyUsage DigitalSignature -KeyExportPolicy Exportable `
            -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(5)
        Write-Host "Created signing certificate $($created.Thumbprint)"
    }
}

if ($Sign) {
    $cert = & $FindCert
    if (-not $cert) { throw 'Signing certificate "CN=SynologyOffice Power BI Connector" not found in Cert:\CurrentUser\My. Run "build.ps1 -NewCert" once to create one.' }
    $makePqx = Join-Path $SdkTools 'MakePQX.exe'
    if (-not (Test-Path $makePqx)) { throw "MakePQX.exe not found in $SdkTools (NuGet package Microsoft.PowerQuery.SdkTools)" }

    $pqx = Join-Path $out 'SynologyOffice.pqx'
    if (Test-Path $pqx) { [IO.File]::Delete($pqx) }
    # MakePQX only accepts a .pfx, so export a throwaway one and delete it right after signing.
    $pfx = Join-Path ([IO.Path]::GetTempPath()) ([Guid]::NewGuid().ToString() + '.pfx')
    $password = [Guid]::NewGuid().ToString('N')
    try {
        Export-PfxCertificate -Cert $cert -FilePath $pfx -Password (ConvertTo-SecureString $password -AsPlainText -Force) | Out-Null
        & $makePqx pack --mez $mez --target $pqx --certificate $pfx --password $password
        if ($LASTEXITCODE -ne 0) { throw "MakePQX pack failed ($LASTEXITCODE)" }
    } finally { if (Test-Path $pfx) { [IO.File]::Delete($pfx) } }
    # --quiet: MakePQX 2.155.2 crashes (assembly binding) while printing the verify result.
    & $makePqx verify $pqx --quiet
    if ($LASTEXITCODE -ne 0) { throw "MakePQX verify failed ($LASTEXITCODE)" }
    Write-Host "Signed $pqx with $($cert.Thumbprint)"
    $artifact = $pqx
}

if ($Install) {
    $target = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Power BI Desktop\Custom Connectors'
    New-Item -ItemType Directory -Force $target | Out-Null
    # Two copies of the same connector (.mez and .pqx) would clash in Power BI.
    $other = Join-Path $target $(if ($artifact -like '*.pqx') { 'SynologyOffice.mez' } else { 'SynologyOffice.pqx' })
    if (Test-Path $other) { [IO.File]::Delete($other) }
    Copy-Item $artifact $target -Force
    Write-Host "Installed $(Split-Path $artifact -Leaf) to $target"
}
