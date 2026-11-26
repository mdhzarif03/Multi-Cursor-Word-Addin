# Multi-Cursor for Word - Windows installer
# Registers a per-user trusted add-in catalog that contains the add-in manifest.
param(
  [Parameter(Mandatory=$true)][string]$Site,
  [switch]$Uninstall
)
$ErrorActionPreference = 'Stop'
$Site      = $Site.TrimEnd('/')
$ShareName = 'MultiCursorAddin'
$Dir       = Join-Path $env:ProgramData 'MultiCursorWordAddin'
$CatalogId = '{c2f3a9d4-6b1e-4f7a-9a58-3d1e0b7a4c21}'
$RegPath   = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogId"
$UncPath   = "\\$env:COMPUTERNAME\$ShareName"

# Share creation needs admin rights; run just that step elevated (one UAC prompt).
function Invoke-Elevated([string]$Command) {
  $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if ($isAdmin) { Invoke-Expression $Command; return }
  $p = Start-Process powershell -Verb RunAs -Wait -PassThru -WindowStyle Hidden `
       -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-Command',$Command
  if ($p.ExitCode -ne 0) { throw 'The administrator step failed or was cancelled.' }
}

if ($Uninstall) {
  Write-Host 'Removing Multi-Cursor...'
  if (Test-Path $RegPath) { Remove-Item $RegPath -Recurse -Force }
  Invoke-Elevated "if (Get-SmbShare -Name '$ShareName' -ErrorAction SilentlyContinue) { Remove-SmbShare -Name '$ShareName' -Force }"
  if (Test-Path $Dir) { Remove-Item $Dir -Recurse -Force }
  Write-Host 'Done. Restart Word.' -ForegroundColor Green
  return
}

Write-Host 'Installing Multi-Cursor for Word...'
New-Item -ItemType Directory -Path $Dir -Force | Out-Null
Invoke-WebRequest "$Site/manifest.xml" -UseBasicParsing -OutFile (Join-Path $Dir 'multi-cursor.xml')

$user = "$env:USERDOMAIN\$env:USERNAME"
Invoke-Elevated "if (-not (Get-SmbShare -Name '$ShareName' -ErrorAction SilentlyContinue)) { New-SmbShare -Name '$ShareName' -Path '$Dir' -ReadAccess '$user' | Out-Null }"

# Register as a trusted catalog for the current user (Flags 1 = show in menu).
New-Item -Path $RegPath -Force | Out-Null
New-ItemProperty -Path $RegPath -Name 'Id'    -Value $CatalogId -PropertyType String -Force | Out-Null
New-ItemProperty -Path $RegPath -Name 'Url'   -Value $UncPath   -PropertyType String -Force | Out-Null
New-ItemProperty -Path $RegPath -Name 'Flags' -Value 1          -PropertyType DWord  -Force | Out-Null

Write-Host ''
Write-Host 'Installed.' -ForegroundColor Green
Write-Host '1. Close and reopen Word.'
Write-Host '2. Insert > My Add-ins > SHARED FOLDER > Multi-Cursor > Add.'
