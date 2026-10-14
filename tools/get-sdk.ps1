# Downloads the Power Query SDK tools (MakePQX.exe for signing, PQTest.exe for tests)
# from nuget.org into the folder build.ps1 and tests\run-tests.ps1 use by default.
param(
    [string]$Version = '2.155.2',
    [string]$Destination = (Join-Path $env:LOCALAPPDATA "PowerQuerySdkTools\$Version")
)

$ErrorActionPreference = 'Stop'
if (Test-Path (Join-Path $Destination 'PQTest.exe')) {
    Write-Host "Power Query SDK $Version is already in $Destination"
    return
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$package = 'microsoft.powerquery.sdktools'
$url = "https://api.nuget.org/v3-flatcontainer/$package/$Version/$package.$Version.nupkg"
$staging = Join-Path ([IO.Path]::GetTempPath()) ("pq-sdk-" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $staging | Out-Null
try {
    $nupkg = Join-Path $staging 'sdk.zip'
    Write-Host "Downloading $url"
    Invoke-WebRequest $url -OutFile $nupkg -UseBasicParsing
    Expand-Archive $nupkg (Join-Path $staging 'pkg')
    New-Item -ItemType Directory -Force (Split-Path $Destination) | Out-Null
    Move-Item (Join-Path $staging 'pkg\tools') $Destination
} finally {
    Remove-Item $staging -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host "Installed Power Query SDK $Version to $Destination"
