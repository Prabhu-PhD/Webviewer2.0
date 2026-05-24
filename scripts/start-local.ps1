[CmdletBinding()]
param(
    [string]$BindHost = "localhost",
    [int]$Port = 3000,
    [switch]$NoHttps
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$defaultPfxPath = Join-Path $projectRoot "certs\localhost.pfx"
$defaultPassphrasePath = Join-Path $projectRoot "certs\localhost.passphrase.txt"

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    throw "Node.js was not found on PATH."
}

$env:HOST = $BindHost
$env:PORT = [string]$Port

if ($NoHttps) {
    Remove-Item Env:SSL_PFX_FILE -ErrorAction SilentlyContinue
    Remove-Item Env:SSL_PFX_PASSPHRASE -ErrorAction SilentlyContinue
    Write-Warning "Starting without HTTPS because -NoHttps was supplied."
} elseif ((Test-Path -LiteralPath $defaultPfxPath) -and (Test-Path -LiteralPath $defaultPassphrasePath)) {
    $env:SSL_PFX_FILE = $defaultPfxPath
    $env:SSL_PFX_PASSPHRASE = (Get-Content -LiteralPath $defaultPassphrasePath -Raw).Trim()
    Write-Host "Using the default localhost development certificate." -ForegroundColor Green
} else {
    Remove-Item Env:SSL_PFX_FILE -ErrorAction SilentlyContinue
    Remove-Item Env:SSL_PFX_PASSPHRASE -ErrorAction SilentlyContinue
    Write-Warning "No default localhost certificate was found in certs/. Falling back to HTTP."
}

& node (Join-Path $projectRoot "scripts\serve.mjs")
