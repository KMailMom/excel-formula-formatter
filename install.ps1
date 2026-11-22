# Formula Formatter for Excel - installer for Excel on Windows (Microsoft 365 / Excel 2021+).
#
# Usage (PowerShell, no admin rights needed):
#   irm https://kmailmom.github.io/excel-formula-formatter/install.ps1 | iex
#
# What it does:
#   1. Downloads manifest.xml to %LOCALAPPDATA%\ExcelFormulaFormatter\manifest.xml
#   2. Registers that manifest for the current user under
#      HKCU\Software\Microsoft\Office\16.0\Wef\Developer (the same mechanism `npm start` uses)
# The add-in code itself is not installed locally; Excel loads it from https://kmailmom.github.io/excel-formula-formatter/.
# Remove it again with:  irm https://kmailmom.github.io/excel-formula-formatter/uninstall.ps1 | iex

$ErrorActionPreference = 'Stop'
$BaseUrl     = 'https://kmailmom.github.io/excel-formula-formatter/'
$AddinId     = '117076a0-2ea1-4038-be03-5425e8b2a84a'
$InstallDir  = Join-Path $env:LOCALAPPDATA 'ExcelFormulaFormatter'
$Manifest    = Join-Path $InstallDir 'manifest.xml'
$RegKey      = 'HKCU:\Software\Microsoft\Office\16.0\Wef\Developer'

Write-Host 'Formula Formatter for Excel - installer' -ForegroundColor Green

if (-not (Test-Path 'HKCU:\Software\Microsoft\Office\16.0')) {
    Write-Warning 'No Office 16.0 (Microsoft 365 / 2016+) settings found for this user. Continuing anyway.'
}

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest -UseBasicParsing -Uri ($BaseUrl + 'manifest.xml') -OutFile $Manifest

# Sanity check: the downloaded file must be our manifest.
[xml]$xml = Get-Content -Raw -Path $Manifest
if ($xml.OfficeApp.Id -ne $AddinId) {
    Remove-Item $Manifest -Force
    throw "Downloaded manifest has an unexpected Id ('$($xml.OfficeApp.Id)'). Aborting."
}

New-Item -Path $RegKey -Force | Out-Null
New-ItemProperty -Path $RegKey -Name $AddinId -Value $Manifest -PropertyType String -Force | Out-Null

Write-Host "Installed version $($xml.OfficeApp.Version)." -ForegroundColor Green
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Write-Host 'Excel is running: close ALL Excel windows and start Excel again.' -ForegroundColor Yellow
} else {
    Write-Host 'Start Excel. A "Formula Formatter" tab appears on the ribbon.'
}
