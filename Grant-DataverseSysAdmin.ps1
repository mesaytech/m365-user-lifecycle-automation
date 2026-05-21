<#
.SYNOPSIS
    Grants the Dataverse System Administrator security role to a user.

.DESCRIPTION
    Uses Az.Accounts for interactive sign-in and calls the Dataverse Web API to
    assign the System Administrator role (scoped to the root business unit) to
    the target user. Handles the Az 14+ SecureString token format.

.PARAMETER EnvironmentUrl
    Dataverse environment URL (e.g., https://org1c2e4139.crm.dynamics.com).

.PARAMETER UserPrincipalName
    Target user's Dataverse domain name / UPN (e.g., mesay@cloudopslabs.onmicrosoft.com).

.EXAMPLE
    .\Grant-DataverseSysAdmin.ps1 `
        -EnvironmentUrl "https://org1c2e4139.crm.dynamics.com" `
        -UserPrincipalName "mesay@cloudopslabs.onmicrosoft.com"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EnvironmentUrl,
    [Parameter(Mandatory)][string]$UserPrincipalName
)

$ErrorActionPreference = 'Stop'
$envUrl = $EnvironmentUrl.TrimEnd('/')

if (-not (Get-Module -ListAvailable Az.Accounts)) {
    Write-Host "Installing Az.Accounts module..." -ForegroundColor Cyan
    Install-Module Az.Accounts -Scope CurrentUser -Force -AllowClobber
}
Import-Module Az.Accounts

if (-not (Get-AzContext)) {
    Write-Host "Signing in to Azure..." -ForegroundColor Cyan
    Connect-AzAccount | Out-Null
}

Write-Host "Acquiring Dataverse token..." -ForegroundColor Cyan
$tokenResp = Get-AzAccessToken -ResourceUrl $envUrl
$token = if ($tokenResp.Token -is [System.Security.SecureString]) {
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenResp.Token)
    try { [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
} else { $tokenResp.Token }

$headers = @{
    Authorization      = "Bearer $token"
    "Content-Type"     = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
    Accept             = "application/json"
}

function Invoke-DV {
    param([string]$Path, [string]$Method = 'GET', $Body = $null)
    $uri = "$envUrl/api/data/v9.2/$Path"
    if ($Body) {
        return Invoke-RestMethod -Uri $uri -Headers $headers -Method $Method -Body ($Body | ConvertTo-Json -Depth 4)
    }
    return Invoke-RestMethod -Uri $uri -Headers $headers -Method $Method
}

Write-Host "Locating user '$UserPrincipalName'..." -ForegroundColor Cyan
$safeUpn = $UserPrincipalName.Replace("'", "''")
$userFilter = [System.Uri]::EscapeDataString("domainname eq '$safeUpn'")
$user = (Invoke-DV "systemusers?`$filter=$userFilter&`$select=systemuserid,fullname,domainname").value | Select-Object -First 1
if (-not $user) { throw "User '$UserPrincipalName' not found in Dataverse" }
Write-Host "  Found: $($user.fullname) (systemuserid=$($user.systemuserid))" -ForegroundColor Gray

Write-Host "Locating root business unit..." -ForegroundColor Cyan
$rootBu = (Invoke-DV "businessunits?`$filter=_parentbusinessunitid_value eq null&`$select=businessunitid,name").value | Select-Object -First 1
Write-Host "  Root BU: $($rootBu.name) (businessunitid=$($rootBu.businessunitid))" -ForegroundColor Gray

Write-Host "Locating System Administrator role..." -ForegroundColor Cyan
$roleFilter = [System.Uri]::EscapeDataString("name eq 'System Administrator' and _businessunitid_value eq $($rootBu.businessunitid)")
$role = (Invoke-DV "roles?`$filter=$roleFilter&`$select=roleid,name").value | Select-Object -First 1
if (-not $role) { throw "Role 'System Administrator' not found in root business unit" }
Write-Host "  Found: $($role.name) (roleid=$($role.roleid))" -ForegroundColor Gray

Write-Host "Assigning role..." -ForegroundColor Cyan
$assignBody = @{ "@odata.id" = "$envUrl/api/data/v9.2/roles($($role.roleid))" }
try {
    Invoke-DV "systemusers($($user.systemuserid))/systemuserroles_association/`$ref" -Method POST -Body $assignBody | Out-Null
    Write-Host "Role assigned." -ForegroundColor Green
} catch {
    $status = $_.Exception.Response.StatusCode.value__
    if ($status -in 412, 400) {
        Write-Host "User likely already has this role (HTTP $status)." -ForegroundColor Yellow
    } else { throw }
}

Write-Host "`nCurrent role assignments for ${UserPrincipalName}:" -ForegroundColor Cyan
(Invoke-DV "systemusers($($user.systemuserid))/systemuserroles_association?`$select=name").value |
    Select-Object name | Format-Table -AutoSize
