<#
.SYNOPSIS
    Adds Microsoft Graph Application permissions to an app registration and grants admin consent.

.DESCRIPTION
    For a client-credentials flow (clientId + secret, no user context), the app needs
    Application-type Graph permissions, not Delegated. This script:
      1. Adds the requested permissions to the app registration's requiredResourceAccess
      2. Creates appRoleAssignments on the app's service principal (the admin consent grant)
    Existing delegated permissions are left untouched. Existing application grants are
    detected and skipped (idempotent).

    The signed-in user must be a tenant Global Administrator (or have equivalent rights
    to manage app permissions and grant tenant-wide admin consent).

.PARAMETER AppId
    The application (client) ID of the target app registration.

.PARAMETER Permissions
    One or more Graph Application permission names to grant (e.g., 'User.ReadWrite.All').

.EXAMPLE
    .\Grant-AppGraphPermissions.ps1 `
        -AppId "536893eb-722e-4e50-850b-a4f78bbf794e" `
        -Permissions @('User.ReadWrite.All', 'Organization.Read.All', 'Group.ReadWrite.All')
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$AppId,
    [Parameter(Mandatory)][string[]]$Permissions
)

$ErrorActionPreference = 'Stop'
$GraphAppId = "00000003-0000-0000-c000-000000000000"

if (-not (Get-Module -ListAvailable Az.Accounts)) {
    Install-Module Az.Accounts -Scope CurrentUser -Force -AllowClobber
}
Import-Module Az.Accounts
if (-not (Get-AzContext)) { Connect-AzAccount | Out-Null }

$tokenResp = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com/"
$token = if ($tokenResp.Token -is [System.Security.SecureString]) {
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenResp.Token)
    try { [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
} else { $tokenResp.Token }

$headers = @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" }

function Invoke-Graph($Path, $Method = 'GET', $Body = $null) {
    $uri = "https://graph.microsoft.com/v1.0/$Path"
    if ($Body) {
        Invoke-RestMethod -Uri $uri -Headers $headers -Method $Method -Body ($Body | ConvertTo-Json -Depth 10)
    } else {
        Invoke-RestMethod -Uri $uri -Headers $headers -Method $Method
    }
}

# Resolve Microsoft Graph service principal (this tenant's object)
Write-Host "Resolving Microsoft Graph service principal..." -ForegroundColor Cyan
$graphSp = (Invoke-Graph "servicePrincipals?`$filter=appId eq '$GraphAppId'&`$select=id,appRoles").value[0]
if (-not $graphSp) { throw "Microsoft Graph service principal not found in tenant" }
Write-Host "  Graph SP objectId: $($graphSp.id)" -ForegroundColor Gray

# Resolve the target app registration and its service principal
Write-Host "Resolving app registration '$AppId'..." -ForegroundColor Cyan
$app = (Invoke-Graph "applications?`$filter=appId eq '$AppId'&`$select=id,displayName,requiredResourceAccess").value[0]
if (-not $app) { throw "App registration with appId '$AppId' not found" }
Write-Host "  App: $($app.displayName) (objectId $($app.id))" -ForegroundColor Gray

$targetSp = (Invoke-Graph "servicePrincipals?`$filter=appId eq '$AppId'&`$select=id,displayName").value[0]
if (-not $targetSp) { throw "Service principal for app '$AppId' not found" }
Write-Host "  Target SP objectId: $($targetSp.id)" -ForegroundColor Gray

# Resolve each requested permission name to its appRole id (from Graph SP's appRoles)
$desiredRoles = @()
foreach ($pn in $Permissions) {
    $role = $graphSp.appRoles | Where-Object { $_.value -eq $pn -and $_.allowedMemberTypes -contains "Application" }
    if (-not $role) { throw "Application permission '$pn' not found on Microsoft Graph SP" }
    $desiredRoles += [PSCustomObject]@{ Name = $pn; Id = $role.id }
    Write-Host "  Permission '$pn' -> $($role.id)" -ForegroundColor Gray
}

# ---------- Step 1: Update requiredResourceAccess on the application ----------
Write-Host "`nUpdating requiredResourceAccess on application..." -ForegroundColor Cyan

# Build pure-hashtable version of current requiredResourceAccess so ConvertTo-Json
# produces the exact schema the Graph API expects.
$rraNew = @()
$graphFound = $false
foreach ($entry in @($app.requiredResourceAccess)) {
    $accessList = @()
    foreach ($a in @($entry.resourceAccess)) {
        $accessList += @{ id = $a.id; type = $a.type }
    }
    if ($entry.resourceAppId -eq $GraphAppId) {
        $graphFound = $true
        $existingIds = @($accessList | ForEach-Object { $_.id })
        foreach ($r in $desiredRoles) {
            if ($existingIds -contains $r.Id) {
                Write-Host "  Already in requiredResourceAccess: $($r.Name)" -ForegroundColor DarkGray
            } else {
                $accessList += @{ id = $r.Id; type = "Role" }
                Write-Host "  Adding to requiredResourceAccess: $($r.Name)" -ForegroundColor Green
            }
        }
    }
    $rraNew += @{ resourceAppId = $entry.resourceAppId; resourceAccess = $accessList }
}
if (-not $graphFound) {
    $accessList = @()
    foreach ($r in $desiredRoles) {
        $accessList += @{ id = $r.Id; type = "Role" }
        Write-Host "  Adding to requiredResourceAccess (new Graph entry): $($r.Name)" -ForegroundColor Green
    }
    $rraNew += @{ resourceAppId = $GraphAppId; resourceAccess = $accessList }
}

$patchBody = @{ requiredResourceAccess = $rraNew }
Invoke-Graph "applications/$($app.id)" -Method PATCH -Body $patchBody | Out-Null
Write-Host "  PATCH applied." -ForegroundColor Green

# ---------- Step 2: Grant admin consent via appRoleAssignment ----------
Write-Host "`nGranting admin consent (appRoleAssignment) per permission..." -ForegroundColor Cyan
$existingGrants = (Invoke-Graph "servicePrincipals/$($targetSp.id)/appRoleAssignments").value
foreach ($r in $desiredRoles) {
    $already = $existingGrants | Where-Object { $_.appRoleId -eq $r.Id -and $_.resourceId -eq $graphSp.id }
    if ($already) {
        Write-Host "  Already granted: $($r.Name)" -ForegroundColor DarkGray
        continue
    }
    $body = @{
        principalId = $targetSp.id
        resourceId  = $graphSp.id
        appRoleId   = $r.Id
    }
    try {
        Invoke-Graph "servicePrincipals/$($targetSp.id)/appRoleAssignments" -Method POST -Body $body | Out-Null
        Write-Host "  Granted: $($r.Name)" -ForegroundColor Green
    } catch {
        $status = $_.Exception.Response.StatusCode.value__
        Write-Host "  FAILED to grant $($r.Name): HTTP $status" -ForegroundColor Red
        try {
            $stream = $_.Exception.Response.GetResponseStream()
            $stream.Position = 0
            (New-Object System.IO.StreamReader($stream)).ReadToEnd() | Write-Host -ForegroundColor DarkRed
        } catch {}
    }
}

# ---------- Step 3: Verify ----------
Write-Host "`nVerifying final state..." -ForegroundColor Cyan
$grants = (Invoke-Graph "servicePrincipals/$($targetSp.id)/appRoleAssignments").value
Write-Host "Application role assignments for $($targetSp.displayName):" -ForegroundColor Yellow
foreach ($g in $grants) {
    $roleName = ($graphSp.appRoles | Where-Object { $_.id -eq $g.appRoleId }).value
    Write-Host "  - $roleName" -ForegroundColor Green
}
Write-Host "`nDone. Token cache: client credentials tokens may take a minute to reflect new permissions." -ForegroundColor Cyan
