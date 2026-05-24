[CmdletBinding()]
param(
    [string]$FriendlyName = "Webviewer2 localhost dev certificate",
    [string]$OutputDirectory = "certs",
    [int]$ValidityYears = 3,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$outputRoot = Join-Path $projectRoot $OutputDirectory
$pfxPath = Join-Path $outputRoot "localhost.pfx"
$cerPath = Join-Path $outputRoot "localhost.cer"
$passphrasePath = Join-Path $outputRoot "localhost.passphrase.txt"

function Remove-ExistingCertificates {
    param([string]$TargetFriendlyName)

    foreach ($storePath in @("Cert:\CurrentUser\My", "Cert:\CurrentUser\Root")) {
        Get-ChildItem -Path $storePath |
            Where-Object { $_.FriendlyName -eq $TargetFriendlyName } |
            ForEach-Object { Remove-Item -LiteralPath $_.PSPath -Force }
    }
}

if ((Test-Path -LiteralPath $pfxPath) -or (Test-Path -LiteralPath $cerPath) -or (Test-Path -LiteralPath $passphrasePath)) {
    if (-not $Force) {
        throw "Certificate output already exists in '$outputRoot'. Re-run with -Force to replace it."
    }
}

New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null

if ($Force) {
    Remove-ExistingCertificates -TargetFriendlyName $FriendlyName
    foreach ($filePath in @($pfxPath, $cerPath, $passphrasePath)) {
        if (Test-Path -LiteralPath $filePath) {
            Remove-Item -LiteralPath $filePath -Force
        }
    }
}

$passphrase = [Guid]::NewGuid().ToString("N")
$securePassphrase = ConvertTo-SecureString -String $passphrase -AsPlainText -Force

$certificate = New-SelfSignedCertificate `
    -DnsName "localhost" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -FriendlyName $FriendlyName `
    -HashAlgorithm "SHA256" `
    -KeyAlgorithm "RSA" `
    -KeyLength 2048 `
    -KeyExportPolicy Exportable `
    -NotAfter (Get-Date).AddYears($ValidityYears)

Export-PfxCertificate -Cert $certificate -FilePath $pfxPath -Password $securePassphrase | Out-Null
Export-Certificate -Cert $certificate -FilePath $cerPath | Out-Null
Import-Certificate -FilePath $cerPath -CertStoreLocation "Cert:\CurrentUser\Root" | Out-Null
Set-Content -LiteralPath $passphrasePath -Value $passphrase -NoNewline

Write-Host "Created and trusted a localhost certificate for Webviewer2." -ForegroundColor Green
Write-Host "PFX: $pfxPath"
Write-Host "CER: $cerPath"
Write-Host "Passphrase: $passphrasePath"
Write-Host ""
Write-Host "Next:"
Write-Host "  powershell -ExecutionPolicy Bypass -File scripts/dev-doctor.ps1"
Write-Host "  powershell -ExecutionPolicy Bypass -File scripts/start-local.ps1"
