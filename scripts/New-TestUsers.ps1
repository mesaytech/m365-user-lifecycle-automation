<#
.SYNOPSIS
Creates test users in the tenant for end-to-end testing of Flow-M365-Lifecycle-Main.

.DESCRIPTION
Idempotent. Re-running the script:
- Skips users that already exist (matched by UPN).
- Re-assigns the license to existing users if -AssignLicense is specified
  and the user currently has no licenses.

Each user is created with:
- Account enabled
- usageLocation set (required before license assignment)
- A unique 14-character password meeting Entra default policy
- forceChangePasswordNextSignIn = true
- Optional license assignment if -SkuPartNumber is provided

The script writes a CSV of UPN + initial password to .\test-users.csv so you
can use the credentials to verify the test mailboxes were provisioned. Treat
that CSV as sensitive, do not commit it. The repo's .gitignore should cover it.

Prerequisites:
- Microsoft.Graph PowerShell module:
    Install-Module Microsoft.Graph -Scope CurrentUser
- Global Admin or User Administrator on the tenant
- License headroom: -Count users will consume -Count licenses of the SKU

.PARAMETER TenantDomain
The tenant's verified primary domain for UPNs.
Example: cloudopslabs.onmicrosoft.com

.PARAMETER Count
Number of test users to create. Default: 5

.PARAMETER UsageLocation
Two-letter ISO country code. Default: US

.PARAMETER SkuPartNumber
Optional. If provided, assigns this license to each user after creation.
Example: ENTERPRISEPREMIUM (E5)

.EXAMPLE
.\New-TestUsers.ps1 -TenantDomain cloudopslabs.onmicrosoft.com -Count 5 -SkuPartNumber ENTERPRISEPREMIUM

.EXAMPLE
.\New-TestUsers.ps1 -TenantDomain cloudopslabs.onmicrosoft.com -Count 3
# Creates 3 users without licenses (you can assign manually later)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantDomain,

    [int]$Count = 5,

    [string]$UsageLocation = 'US',

    [string]$SkuPartNumber
)

$ErrorActionPreference = 'Stop'

# --- 0. Module check ------------------------------------------------------

$requiredModules = @('Microsoft.Graph.Users', 'Microsoft.Graph.Identity.DirectoryManagement')
foreach ($m in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        throw "$m is not installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
    }
}

# --- 1. Test user roster --------------------------------------------------
# Realistic but obviously-test names. Adjust or extend as needed.
# The script will use the first $Count entries.

$roster = @(
    @{ First = 'Alex';   Last = 'Morgan'; Title = 'Software Engineer';       Department = 'Engineering'; OfficeLocation = 'Remote' }
    @{ First = 'Jamie';  Last = 'Chen';   Title = 'Product Manager';         Department = 'Engineering'; OfficeLocation = 'HQ' }
    @{ First = 'Sam';    Last = 'Patel';  Title = 'Systems Administrator';   Department = 'IT';          OfficeLocation = 'HQ' }
    @{ First = 'Riley';  Last = 'Kim';    Title = 'Finance Analyst';         Department = 'Finance';     OfficeLocation = 'HQ' }
    @{ First = 'Jordan'; Last = 'Reyes';  Title = 'HR Coordinator';          Department = 'HR';          OfficeLocation = 'Remote' }
    @{ First = 'Taylor'; Last = 'Singh';  Title = 'Operations Specialist';   Department = 'Operations';  OfficeLocation = 'Field' }
    @{ First = 'Casey';  Last = 'Nguyen'; Title = 'Sales Representative';    Department = 'Sales';       OfficeLocation = 'Remote' }
    @{ First = 'Morgan'; Last = 'Becker'; Title = 'Executive Assistant';     Department = 'Executive';   OfficeLocation = 'HQ' }
)

if ($Count -gt $roster.Count) {
    throw "Requested $Count users but the roster only has $($roster.Count) entries. Extend the roster array in the script."
}

$selectedRoster = $roster[0..($Count - 1)]

# --- 2. Connect to Graph --------------------------------------------------

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
$scopes = @('User.ReadWrite.All', 'Directory.ReadWrite.All', 'Organization.Read.All')
Connect-MgGraph -Scopes $scopes -NoWelcome | Out-Null

# --- 3. Resolve SKU if license assignment requested -----------------------

$skuId = $null
if ($SkuPartNumber) {
    Write-Host "Looking up SkuId for $SkuPartNumber..." -ForegroundColor Cyan
    $sku = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $SkuPartNumber }
    if (-not $sku) {
        throw "SKU '$SkuPartNumber' not found in this tenant. Run Get-MgSubscribedSku to list available SKUs."
    }
    $skuId = $sku.SkuId
    $available = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
    Write-Host "  SkuId    : $skuId"
    Write-Host "  Available: $available of $($sku.PrepaidUnits.Enabled)"
    if ($available -lt $Count) {
        Write-Warning "Only $available licenses available, but $Count users will be created. Some will fail license assignment."
    }
}

# --- 4. Create users ------------------------------------------------------

$results = @()

foreach ($entry in $selectedRoster) {
    $first = $entry.First
    $last = $entry.Last
    $upn = "$($first.ToLower()).$($last.ToLower())@$TenantDomain"
    $mailNickname = "$($first.ToLower()).$($last.ToLower())"
    $displayName = "$first $last"

    Write-Host ""
    Write-Host "User: $displayName ($upn)" -ForegroundColor Cyan

    # Existing user check
    $existing = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  [skip] User already exists, id=$($existing.Id)" -ForegroundColor Yellow

        # Optionally back-fill license
        if ($skuId) {
            $current = (Get-MgUser -UserId $existing.Id -Property AssignedLicenses).AssignedLicenses
            if (-not ($current | Where-Object { $_.SkuId -eq $skuId })) {
                Write-Host "  [lic ] assigning $SkuPartNumber to existing user..." -ForegroundColor Green
                Set-MgUserLicense -UserId $existing.Id -AddLicenses @(@{ SkuId = $skuId }) -RemoveLicenses @() | Out-Null
            }
        }

        $results += [PSCustomObject]@{
            DisplayName    = $displayName
            UPN            = $upn
            ObjectId       = $existing.Id
            InitialPassword = '(existing user, password not reset)'
            LicenseAssigned = if ($skuId) { $SkuPartNumber } else { '(none)' }
            Status         = 'Existed'
        }
        continue
    }

    # Generate per-user password using the same pattern as the flow's varTempPassword
    $pw = (([guid]::NewGuid().ToString() -replace '-', '').Substring(0, 11)) + 'A7!'

    $passwordProfile = @{
        Password                      = $pw
        ForceChangePasswordNextSignIn = $true
    }

    $newUserParams = @{
        DisplayName       = $displayName
        UserPrincipalName = $upn
        MailNickname      = $mailNickname
        AccountEnabled    = $true
        UsageLocation     = $UsageLocation
        PasswordProfile   = $passwordProfile
        JobTitle          = $entry.Title
        Department        = $entry.Department
        OfficeLocation    = $entry.OfficeLocation
    }

    Write-Host "  [add ] creating user..." -ForegroundColor Green
    $newUser = New-MgUser @newUserParams

    $assignedSku = '(none)'
    if ($skuId) {
        try {
            Write-Host "  [lic ] assigning $SkuPartNumber..." -ForegroundColor Green
            Set-MgUserLicense -UserId $newUser.Id -AddLicenses @(@{ SkuId = $skuId }) -RemoveLicenses @() | Out-Null
            $assignedSku = $SkuPartNumber
        }
        catch {
            Write-Warning "  License assignment failed for ${upn}: $($_.Exception.Message)"
        }
    }

    $results += [PSCustomObject]@{
        DisplayName    = $displayName
        UPN            = $upn
        ObjectId       = $newUser.Id
        InitialPassword = $pw
        LicenseAssigned = $assignedSku
        Status         = 'Created'
    }
}

# --- 5. Output ------------------------------------------------------------

Write-Host ""
Write-Host "=== Summary ===" -ForegroundColor Green
$results | Format-Table DisplayName, UPN, Status, LicenseAssigned -AutoSize

# CSV with credentials, written next to the script
$csvPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'test-users.csv'
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Credentials CSV written to: $csvPath" -ForegroundColor Yellow
Write-Host "This file contains initial passwords. Do not commit it. .gitignore should exclude test-users.csv."
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Wait 2-5 minutes for license-driven mailbox provisioning to complete."
Write-Host "  2. Pick one user from the roster as the target for the Convert-MailboxToShared runbook test."
Write-Host "  3. Use other roster entries as test targets for the flow's New Hire and Termination tests in Task 6."

Disconnect-MgGraph | Out-Null
