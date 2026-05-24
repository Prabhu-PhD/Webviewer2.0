param(
    [string[]]$Paths = @("manifest.localhost.xml", "manifest.hosted.xml")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

foreach ($Path in $Paths) {
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Manifest not found: $Path"
    }

    [xml]$Manifest = Get-Content -LiteralPath $Path -Raw
    $OfficeApp = $Manifest.OfficeApp

    Assert-True ($null -ne $OfficeApp) "Missing OfficeApp root in $Path"

    $Type = $OfficeApp.GetAttribute("type", "http://www.w3.org/2001/XMLSchema-instance")
    Assert-True ($Type -eq "ContentApp") "Expected xsi:type='ContentApp' in $Path"

    $DisplayName = $OfficeApp.DisplayName.DefaultValue
    Assert-True (-not [string]::IsNullOrWhiteSpace($DisplayName)) "DisplayName is missing in $Path"

    $SourceLocation = $OfficeApp.DefaultSettings.SourceLocation.DefaultValue
    $RequestedHeight = [int]$OfficeApp.DefaultSettings.RequestedHeight
    $RequestedWidth = [int]$OfficeApp.DefaultSettings.RequestedWidth
    $Permissions = [string]$OfficeApp.Permissions
    $HostName = [string]$OfficeApp.Hosts.Host.Name

    $SourceUri = $null
    Assert-True ([System.Uri]::TryCreate($SourceLocation, [System.UriKind]::Absolute, [ref]$SourceUri)) "SourceLocation is not an absolute URL in $Path"
    Assert-True ($SourceUri.Scheme -eq "https") "SourceLocation must use HTTPS in $Path"
    Assert-True ($RequestedHeight -ge 32 -and $RequestedHeight -le 1000) "RequestedHeight must be between 32 and 1000 in $Path"
    Assert-True ($RequestedWidth -ge 32 -and $RequestedWidth -le 1000) "RequestedWidth must be between 32 and 1000 in $Path"
    Assert-True ($Permissions -eq "ReadWriteDocument") "Permissions must be ReadWriteDocument in $Path"
    Assert-True ($HostName -eq "Presentation") "Host must be Presentation in $Path"

    Write-Host "Validated $Path" -ForegroundColor Green
}
