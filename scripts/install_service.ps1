<#
Install helper for DocCompareService using NSSM.

Usage (Run PowerShell as Administrator):
  .\install_service.ps1 -ServiceName DocCompareService -PythonExe "C:\Path\to\python.exe" -Module "src.doc_compare.service"

This script downloads NSSM, extracts the appropriate nssm.exe, and registers the service.
It does not require installing pywin32 service host files.
#>

param(
    [string]$ServiceName = "DocCompareService",
    [string]$PythonExe = "C:\Users\3874\AppData\Local\Programs\Python\Python313\python.exe",
    [string]$Module = "src.doc_compare.service",
    [string]$NssmUrl = "https://nssm.cc/release/nssm-2.24.zip",
    [switch]$InstallToSystem32
)

function Assert-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Error "This script must be run as Administrator. Right-click PowerShell and choose 'Run as Administrator'."
        exit 1
    }
}

Assert-Admin

$tmp = Join-Path $env:TEMP ([System.Guid]::NewGuid().ToString())
New-Item -Path $tmp -ItemType Directory -Force | Out-Null
$zip = Join-Path $tmp "nssm.zip"

Write-Host "Downloading NSSM..."
Invoke-WebRequest -Uri $NssmUrl -OutFile $zip -UseBasicParsing

Write-Host "Extracting..."
Expand-Archive -Path $zip -DestinationPath $tmp -Force

$arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') { 'win64' } else { 'win32' }
$nssmCandidate = Get-ChildItem -Path $tmp -Recurse -Filter 'nssm.exe' | Where-Object { $_.FullName -like "*\$arch\*" } | Select-Object -First 1
if (-not $nssmCandidate) {
    $nssmCandidate = Get-ChildItem -Path $tmp -Recurse -Filter 'nssm.exe' | Select-Object -First 1
}

if (-not $nssmCandidate) {
    Write-Error "Could not find nssm.exe in archive. Please download NSSM manually from https://nssm.cc/download"
    exit 1
}

$nssmPath = $nssmCandidate.FullName

if ($InstallToSystem32) {
    $dest = Join-Path $env:SystemRoot 'System32\nssm.exe'
    Copy-Item -Path $nssmPath -Destination $dest -Force
    $nssmPath = $dest
    Write-Host "Copied nssm.exe to $dest"
} else {
    Write-Host "Using nssm from temporary path: $nssmPath"
}

$binArgs = "-m $Module"

Write-Host "Installing service '$ServiceName' -> $PythonExe $binArgs"
& $nssmPath install $ServiceName $PythonExe $binArgs

if ($LASTEXITCODE -ne 0) {
    Write-Error "nssm install failed with exit code $LASTEXITCODE"
    exit 1
}

Write-Host "Setting service to start automatically"
& $nssmPath set $ServiceName Start SERVICE_AUTO_START

Write-Host "Starting service"
& $nssmPath start $ServiceName

Write-Host "Cleaning up temporary files"
Remove-Item -Path $tmp -Recurse -Force

Write-Host "Done. Service '$ServiceName' should be installed and started (if no errors)."
