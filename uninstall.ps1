# Formula Formatter for Excel - uninstaller for Excel on Windows.
# Usage:  irm https://kmailmom.github.io/excel-formula-formatter/uninstall.ps1 | iex

$ErrorActionPreference = 'Stop'
$AddinId    = '117076a0-2ea1-4038-be03-5425e8b2a84a'
$InstallDir = Join-Path $env:LOCALAPPDATA 'ExcelFormulaFormatter'
$RegKey     = 'HKCU:\Software\Microsoft\Office\16.0\Wef\Developer'

if (Test-Path $RegKey) {
    Remove-ItemProperty -Path $RegKey -Name $AddinId -ErrorAction SilentlyContinue
}
if (Test-Path $InstallDir) { Remove-Item -Recurse -Force $InstallDir }

Write-Host 'Formula Formatter for Excel removed. Restart Excel to make the ribbon tab disappear.' -ForegroundColor Green
