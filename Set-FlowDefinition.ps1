<#
.SYNOPSIS
    Uploads an edited Power Automate flow definition via PATCH to the Power Automate REST API.

.DESCRIPTION
    Reads a JSON file produced by Get-FlowDefinition.ps1 (or hand-edited equivalent)
    and PATCHes the flow's properties.definition back to the environment. Only the
    definition is updated — runtime state, owner, connections, etc. are preserved.

.PARAMETER EnvironmentId
    BAP environment ID (e.g., Default-f2ce44ac-e0cd-4b5a-a300-5fb2e149d210).

.PARAMETER FlowName
    Flow GUID (e.g., abc12345-...) or display name. GUID is preferred.

.PARAMETER DefinitionPath
    Path to the local JSON file containing the edited flow definition. Must match
    the structure produced by Get-FlowDefinition.ps1 (a full flow envelope with
    properties.definition somewhere inside).

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Set-FlowDefinition.ps1 `
        -EnvironmentId "Default-f2ce44ac-e0cd-4b5a-a300-5fb2e149d210" `
        -FlowName "Flow-M365-Lifecycle-Main" `
        -DefinitionPath ".\flow-Flow-M365-Lifecycle-Main.json"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EnvironmentId,
    [Parameter(Mandatory)][string]$FlowName,
    [Parameter(Mandatory)][string]$DefinitionPath,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $DefinitionPath)) {
    throw "Definition file not found: $DefinitionPath"
}

if (-not (Get-Module -ListAvailable Az.Accounts)) {
    Install-Module Az.Accounts -Scope CurrentUser -Force -AllowClobber
}
Import-Module Az.Accounts

if (-not (Get-AzContext)) {
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
} else {
    Write-Host "Resolving flow '$FlowName' to GUID..." -ForegroundColor Cyan
    $listUri = "${apiBase}?api-version=$apiVersion"
    $flows = (Invoke-RestMethod -Uri $listUri -Headers $headers -Method GET).value
    $match = $flows | Where-Object { $_.properties.displayName -eq $FlowName }
    if (-not $match) { throw "Flow '$FlowName' not found in environment $EnvironmentId" }
    if ($match.Count -gt 1) { throw "Multiple flows match '$FlowName' - use the GUID instead" }
    $flowId = $match.name
}

# Load and parse the definition file
Write-Host "Loading definition from $DefinitionPath..." -ForegroundColor Cyan
$raw = Get-Content -Path $DefinitionPath -Raw -Encoding utf8
$parsed = $raw | ConvertFrom-Json -Depth 100
if (-not $parsed.properties -or -not $parsed.properties.definition) {
    throw "File does not contain properties.definition. Expected structure from Get-FlowDefinition.ps1."
}

$displayName = $parsed.properties.displayName
$actionCount = if ($parsed.properties.definition.actions) {
    ($parsed.properties.definition.actions.PSObject.Properties | Measure-Object).Count
} else { 0 }

Write-Host "`nReady to PATCH:" -ForegroundColor Yellow
Write-Host "  Flow:            $displayName" -ForegroundColor Gray
Write-Host "  GUID:            $flowId" -ForegroundColor Gray
Write-Host "  Top-level acts:  $actionCount" -ForegroundColor Gray
Write-Host "  Env:             $EnvironmentId" -ForegroundColor Gray

if (-not $Force) {
    $confirm = Read-Host "`nProceed with PATCH? (y/n)"
    if ($confirm -ne 'y' -and $confirm -ne 'Y') {
        Write-Host "Aborted." -ForegroundColor Yellow
        exit 0
    }
}

# Build PATCH body — only updates properties.definition, leaves everything else untouched
$patchBody = @{
    properties = @{
        definition = $parsed.properties.definition
    }
} | ConvertTo-Json -Depth 100

# Optionally include connectionReferences if present (so existing connections stay wired)
if ($parsed.properties.connectionReferences) {
    $patchObj = $patchBody | ConvertFrom-Json -Depth 100
    $patchObj.properties | Add-Member -NotePropertyName connectionReferences -NotePropertyValue $parsed.properties.connectionReferences -Force
    $patchBody = $patchObj | ConvertTo-Json -Depth 100
}

$flowUri = "$apiBase/${flowId}?api-version=$apiVersion"
Write-Host "`nPATCH $flowUri" -ForegroundColor Cyan

try {
    $response = Invoke-RestMethod -Uri $flowUri -Headers $headers -Method PATCH -Body $patchBody
    Write-Host "`nFlow definition updated." -ForegroundColor Green
    Write-Host "  New state: $($response.properties.state)" -ForegroundColor Gray
    Write-Host "`nVerify in Power Automate UI — refresh the flow page." -ForegroundColor Cyan
} catch {
    $status = $_.Exception.Response.StatusCode.value__
    $body = ""
    if ($_.Exception.Response) {
        try {
            $stream = $_.Exception.Response.GetResponseStream()
            $stream.Position = 0
            $body = (New-Object System.IO.StreamReader($stream)).ReadToEnd()
        } catch {}
    }
    Write-Host "`nPATCH failed: HTTP $status" -ForegroundColor Red
    if ($body) { Write-Host $body -ForegroundColor Red }
    throw
}
