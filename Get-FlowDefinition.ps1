<#
.SYNOPSIS
    Downloads a Power Automate flow's full definition as JSON for local editing.

.DESCRIPTION
    Calls the Power Automate Management REST API to GET a flow by name or GUID,
    then writes the response JSON to a local file. Use Set-FlowDefinition.ps1 to
    PATCH the edited file back. Requires only flow ownership (no Dataverse System
    Administrator role needed).

.PARAMETER EnvironmentId
    BAP environment ID (e.g., Default-f2ce44ac-e0cd-4b5a-a300-5fb2e149d210).

.PARAMETER FlowName
    Either the flow's GUID (e.g., abc12345-...) or its display name
    (e.g., Flow-M365-Lifecycle-Main). Display name is resolved to GUID
    by listing flows in the environment.

.PARAMETER OutputPath
    Local path to write the JSON file. Defaults to .\flow-{name}.json.

.EXAMPLE
    .\Get-FlowDefinition.ps1 `
        -EnvironmentId "Default-f2ce44ac-e0cd-4b5a-a300-5fb2e149d210" `
        -FlowName "Flow-M365-Lifecycle-Main"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EnvironmentId,
    [Parameter(Mandatory)][string]$FlowName,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable Az.Accounts)) {
    Write-Host "Installing Az.Accounts module..." -ForegroundColor Cyan
    Install-Module Az.Accounts -Scope CurrentUser -Force -AllowClobber
}
Import-Module Az.Accounts

if (-not (Get-AzContext)) {
    Write-Host "Signing in to Azure..." -ForegroundColor Cyan
    Connect-AzAccount | Out-Null
}

Write-Host "Acquiring token for Power Automate API..." -ForegroundColor Cyan
$tokenResp = Get-AzAccessToken -ResourceUrl "https://service.flow.microsoft.com/"
$token = if ($tokenResp.Token -is [System.Security.SecureString]) {
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenResp.Token)
    try { [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
} else { $tokenResp.Token }

$headers = @{
    Authorization  = "Bearer $token"
    "Content-Type" = "application/json"
}

$apiBase = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/flows"
$apiVersion = "2016-11-01"

# Resolve flow name → GUID if needed
if ($FlowName -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
    $flowId = $FlowName
    Write-Host "Using flow GUID directly: $flowId" -ForegroundColor Gray
} else {
    Write-Host "Resolving flow '$FlowName' to GUID..." -ForegroundColor Cyan
    $listUri = "${apiBase}?api-version=$apiVersion"
    $flows = (Invoke-RestMethod -Uri $listUri -Headers $headers -Method GET).value
    $match = $flows | Where-Object { $_.properties.displayName -eq $FlowName }
    if (-not $match) {
        Write-Host "Available flows:" -ForegroundColor Yellow
        $flows | ForEach-Object { Write-Host "  - $($_.properties.displayName) ($($_.name))" -ForegroundColor Gray }
        throw "Flow '$FlowName' not found in environment $EnvironmentId"
    }
    if ($match.Count -gt 1) {
        throw "Multiple flows match '$FlowName' - use the GUID instead"
    }
    $flowId = $match.name
    Write-Host "  Resolved to GUID: $flowId" -ForegroundColor Gray
}

# GET the flow definition
Write-Host "Fetching flow definition..." -ForegroundColor Cyan
$flowUri = "$apiBase/${flowId}?api-version=$apiVersion"
$response = Invoke-RestMethod -Uri $flowUri -Headers $headers -Method GET

if (-not $OutputPath) {
    $safeName = ($response.properties.displayName -replace '[^a-zA-Z0-9_-]', '_')
    $OutputPath = ".\flow-$safeName.json"
}

$json = $response | ConvertTo-Json -Depth 100
$json | Out-File -FilePath $OutputPath -Encoding utf8

$sizeKB = [Math]::Round((Get-Item $OutputPath).Length / 1024, 1)
Write-Host "`nSaved flow definition:" -ForegroundColor Green
Write-Host "  File:        $((Resolve-Path $OutputPath).Path)" -ForegroundColor Gray
Write-Host "  Size:        $sizeKB KB" -ForegroundColor Gray
Write-Host "  Flow:        $($response.properties.displayName)" -ForegroundColor Gray
Write-Host "  GUID:        $flowId" -ForegroundColor Gray
Write-Host "  State:       $($response.properties.state)" -ForegroundColor Gray
Write-Host "`nEdit properties.definition in the JSON file, then push back with:" -ForegroundColor Cyan
Write-Host "  .\Set-FlowDefinition.ps1 -EnvironmentId '$EnvironmentId' -FlowName '$flowId' -DefinitionPath '$OutputPath'" -ForegroundColor Gray
