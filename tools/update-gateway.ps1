# Run as Administrator on the gateway machine. Puts the freshly built connector into the
# on-premises data gateway's custom connectors folder and restarts the gateway service.
# The folder is read from the gateway config (MashupCustomConnectorsDirectory). Any
# SynologyOffice.mez/.pqx already there is backed up to bin\gateway-backup first, and the
# other variant is removed so the gateway does not load two copies of the connector.
param(
    [string]$Source = (Join-Path $PSScriptRoot '..\bin\SynologyOffice.mez'),
    [string]$GatewayDir = 'C:\Program Files\On-premises data gateway'
)

$ErrorActionPreference = 'Stop'
$identity = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $identity.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run this script in PowerShell started "As Administrator".'
}
$Source = (Resolve-Path $Source).Path

$config = [xml](Get-Content (Join-Path $GatewayDir 'Microsoft.PowerBI.DataMovement.Pipeline.GatewayCore.dll.config') -Raw)
$setting = $config.SelectSingleNode("//setting[@name='MashupCustomConnectorsDirectory']/value")
if (-not $setting -or -not $setting.InnerText) { throw 'MashupCustomConnectorsDirectory is not set in the gateway config; set the custom connectors folder in the gateway app first.' }
$target = $setting.InnerText
New-Item -ItemType Directory -Force $target | Out-Null

$backup = Join-Path $PSScriptRoot '..\bin\gateway-backup'
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
foreach ($old in Get-ChildItem $target -Filter 'SynologyOffice.*' | Where-Object { $_.Extension -in '.mez', '.pqx' }) {
    New-Item -ItemType Directory -Force $backup | Out-Null
    Copy-Item $old.FullName (Join-Path $backup "$($old.BaseName).$stamp$($old.Extension)") -Force
    if ($old.Extension -ne [IO.Path]::GetExtension($Source)) { Remove-Item $old.FullName -Force }
    Write-Host "Backed up $($old.Name)"
}

Copy-Item $Source $target -Force
Write-Host "Copied $(Split-Path $Source -Leaf) to $target"
Restart-Service PBIEgwService
$installed = Join-Path $target (Split-Path $Source -Leaf)
$same = (Get-FileHash $installed).Hash -eq (Get-FileHash $Source).Hash
Write-Host "Matches build: $same; gateway service: $((Get-Service PBIEgwService).Status)"
