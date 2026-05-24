[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$pfxPath = Join-Path $projectRoot "certs\localhost.pfx"
$passphrasePath = Join-Path $projectRoot "certs\localhost.passphrase.txt"
$manifestPath = Join-Path $projectRoot "manifest.localhost.xml"

Write-Host "Webviewer2 local doctor" -ForegroundColor Cyan
Write-Host ""

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    throw "Node.js was not found on PATH."
}

$nodeVersion = (& node -v).Trim()
Write-Host "Node: $nodeVersion"

if (-not (Test-Path -LiteralPath $manifestPath)) {
    throw "Missing manifest.localhost.xml"
}

& (Join-Path $PSScriptRoot "validate-manifest.ps1")

if ((Test-Path -LiteralPath $pfxPath) -and (Test-Path -LiteralPath $passphrasePath)) {
    Write-Host "HTTPS cert: ready" -ForegroundColor Green
    Write-Host "  $pfxPath"
} else {
    Write-Warning "HTTPS cert: missing. Run scripts/new-dev-cert.ps1 before testing in Office."
}

Write-Host ""
Write-Host "If PowerPoint desktop says 'We can't open this add-in from localhost', use an elevated prompt and run:"
Write-Host '  CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"'
Write-Host ""
Write-Host "Next:"
Write-Host "  powershell -ExecutionPolicy Bypass -File scripts/start-local.ps1"
