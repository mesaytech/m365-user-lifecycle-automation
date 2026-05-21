<#
.SYNOPSIS
Creates the SharePoint Communication Site and LifecycleAuditLog list described in
docs/01-sharepoint-schema.md.

.DESCRIPTION
Idempotent. Re-running this script:
- Skips site creation if the site already exists.
- Skips list creation if the list already exists.
- Adds missing columns. Existing columns are left untouched (PnP throws on duplicates,
  which the script catches and treats as "already there").
- Re-applies indexes and view definitions every run.

Prerequisites:
- PowerShell 7 (recommended) or Windows PowerShell 5.1.
- PnP.PowerShell module installed:
    Install-Module PnP.PowerShell -Scope CurrentUser
- An Entra ID app registration for PnP PowerShell, separate from the
  M365-Lifecycle-Automation app used by the Power Automate flow. PnP.PowerShell 2.x
  removed the bundled multi-tenant Management Shell app, so you bring your own.
  Fastest path:
    Register-PnPEntraIDApp -ApplicationName "PnP-Lifecycle-Setup" -Tenant <tenant>.onmicrosoft.com -Interactive
  This creates the app and prompts for admin consent. Record the ClientId it returns
  and pass it to this script via -ClientId. The PNPPOWERSHELL_CLIENTID environment
  variable is not honored by PnP.PowerShell 2.x in current builds, hence the explicit
  parameter.
- Global Admin or SharePoint Admin on the tenant.

.PARAMETER TenantHostname
The tenant's SharePoint hostname, without protocol or path.
Example: contoso.sharepoint.com

.PARAMETER ClientId
GUID of the Entra ID app registration created for PnP PowerShell. Passed to every
internal Connect-PnPOnline call.

.PARAMETER SiteAlias
The URL segment after /sites/ for the new site. Default: ITAutomation

.PARAMETER SiteTitle
The display name of the new site. Default: IT Automation

.PARAMETER ListName
The internal list name. Default: LifecycleAuditLog

.EXAMPLE
.\Setup-SharePointList.ps1 -TenantHostname contoso.sharepoint.com -ClientId "9ce48a27-ef22-4210-b98c-7bedf274cf09"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantHostname,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$ClientId,

    [string]$SiteAlias = 'ITAutomation',

    [string]$SiteTitle = 'IT Automation',

    [string]$ListName = 'LifecycleAuditLog'
)

$ErrorActionPreference = 'Stop'

# --- 0. Module check ------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
}

Import-Module PnP.PowerShell -DisableNameChecking

# --- 1. Resolve URLs ------------------------------------------------------

$TenantHostname = $TenantHostname.TrimEnd('/').Replace('https://', '').Replace('http://', '')
$AdminUrl = "https://$($TenantHostname.Split('.')[0])-admin.sharepoint.com"
$TenantUrl = "https://$TenantHostname"
$SiteUrl = "$TenantUrl/sites/$SiteAlias"

Write-Host "Tenant URL  : $TenantUrl"
Write-Host "Admin URL   : $AdminUrl"
Write-Host "Site URL    : $SiteUrl"
Write-Host "List name   : $ListName"
Write-Host "Client ID   : $ClientId"
Write-Host ""

# --- 2. Create site if missing -------------------------------------------

Write-Host "Connecting to tenant admin..." -ForegroundColor Cyan
Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId

$existingSite = Get-PnPTenantSite -Identity $SiteUrl -ErrorAction SilentlyContinue

if ($existingSite) {
    Write-Host "Site already exists at $SiteUrl, skipping create." -ForegroundColor Yellow
}
else {
    Write-Host "Creating Communication Site $SiteUrl..." -ForegroundColor Cyan
    New-PnPSite `
        -Type CommunicationSite `
        -Title $SiteTitle `
        -Url $SiteUrl `
        -Description "Hosts the LifecycleAuditLog list used by Flow-M365-Lifecycle-Main." | Out-Null
    Start-Sleep -Seconds 10
}

# --- 3. Connect to the new site ------------------------------------------

Write-Host "Connecting to $SiteUrl..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# --- 4. Create the list ---------------------------------------------------

$list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue

if ($list) {
    Write-Host "List $ListName already exists, skipping create." -ForegroundColor Yellow
}
else {
    Write-Host "Creating list $ListName..." -ForegroundColor Cyan
    $list = New-PnPList -Title $ListName -Template GenericList -EnableVersioning -OnQuickLaunch
    Set-PnPList -Identity $ListName -MajorVersions 50
}

# Rename the default Title column display name to "Audit Record ID"
Set-PnPField -List $ListName -Identity 'Title' -Values @{ Title = 'Audit Record ID' } | Out-Null

# --- 5. Column schema -----------------------------------------------------
# Defined as an ordered array so the columns appear in the list in a predictable order.

$columns = @(
    @{ InternalName = 'ActionType';              DisplayName = 'Action Type';                   Type = 'Choice';   Required = $true;  Choices = @('New Hire', 'Termination') }
    @{ InternalName = 'TargetUserUPN';           DisplayName = 'Target User UPN';               Type = 'Text';     Required = $true }
    @{ InternalName = 'TargetUserDisplayName';   DisplayName = 'Target User Display Name';      Type = 'Text';     Required = $false }
    @{ InternalName = 'TargetUserObjectID';      DisplayName = 'Target User Object ID';         Type = 'Text';     Required = $false }
    @{ InternalName = 'SubmitterEmail';          DisplayName = 'Submitter Email';               Type = 'Text';     Required = $true }
    @{ InternalName = 'SubmitterDisplayName';    DisplayName = 'Submitter Display Name';        Type = 'Text';     Required = $false }
    @{ InternalName = 'FormResponseID';          DisplayName = 'Form Response ID';              Type = 'Text';     Required = $true }
    @{ InternalName = 'FlowRunID';               DisplayName = 'Flow Run ID';                   Type = 'Text';     Required = $true }
    @{ InternalName = 'FlowRunURL';              DisplayName = 'Flow Run URL';                  Type = 'URL';      Required = $false }
    @{ InternalName = 'StartTimestamp';          DisplayName = 'Start Timestamp';               Type = 'DateTime'; Required = $true;  DisplayFormat = 1 }
    @{ InternalName = 'EndTimestamp';            DisplayName = 'End Timestamp';                 Type = 'DateTime'; Required = $false; DisplayFormat = 1 }
    @{ InternalName = 'DurationSeconds';         DisplayName = 'Duration (seconds)';            Type = 'Number';   Required = $false }
    @{ InternalName = 'Status';                  DisplayName = 'Status';                        Type = 'Choice';   Required = $true;  Choices = @('In Progress', 'Succeeded', 'Partial Success', 'Failed') }
    @{ InternalName = 'StepsCompleted';          DisplayName = 'Steps Completed';               Type = 'Note';     Required = $false; PlainText = $true }
    @{ InternalName = 'StepsFailed';             DisplayName = 'Steps Failed';                  Type = 'Note';     Required = $false; PlainText = $true }
    @{ InternalName = 'ErrorDetails';            DisplayName = 'Error Details';                 Type = 'Note';     Required = $false; PlainText = $true }
    @{ InternalName = 'RetryCount';              DisplayName = 'Retry Count';                   Type = 'Number';   Required = $false }
    @{ InternalName = 'LicenseSKU';              DisplayName = 'License SKU';                   Type = 'Text';     Required = $false }
    @{ InternalName = 'LicenseSKUID';            DisplayName = 'License SKU ID';                Type = 'Text';     Required = $false }
    @{ InternalName = 'Department';              DisplayName = 'Department';                    Type = 'Choice';   Required = $false; Choices = @('(none)', 'IT', 'Finance', 'HR', 'Operations', 'Engineering', 'Sales', 'Executive') }
    @{ InternalName = 'ManagerUPN';              DisplayName = 'Manager UPN';                   Type = 'Text';     Required = $false }
    @{ InternalName = 'StartDate';               DisplayName = 'Start Date';                    Type = 'DateTime'; Required = $false; DisplayFormat = 0 }
    @{ InternalName = 'TerminationDate';         DisplayName = 'Termination Date';              Type = 'DateTime'; Required = $false; DisplayFormat = 1 }
    @{ InternalName = 'MailboxConverted';        DisplayName = 'Mailbox Converted to Shared';   Type = 'Boolean';  Required = $false }
    @{ InternalName = 'MailboxForwardingTo';     DisplayName = 'Mailbox Forwarding To';         Type = 'Text';     Required = $false }
    @{ InternalName = 'GroupsAdded';             DisplayName = 'Groups Added';                  Type = 'Note';     Required = $false; PlainText = $true }
    @{ InternalName = 'GroupsRemoved';           DisplayName = 'Groups Removed';                Type = 'Note';     Required = $false; PlainText = $true }
    @{ InternalName = 'TerminationReason';       DisplayName = 'Termination Reason';            Type = 'Choice';   Required = $false; Choices = @('(none)', 'Voluntary resignation', 'Involuntary', 'Retirement', 'Contract end', 'Other') }
    @{ InternalName = 'RetentionPeriod';         DisplayName = 'Retention Period';              Type = 'Choice';   Required = $false; Choices = @('(none)', '30 days', '60 days', '90 days', 'Indefinite') }
    @{ InternalName = 'Notes';                   DisplayName = 'Notes';                         Type = 'Note';     Required = $false; PlainText = $false }
)

# --- 6. Add columns -------------------------------------------------------

foreach ($col in $columns) {
    $existing = Get-PnPField -List $ListName -Identity $col.InternalName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  [skip] $($col.InternalName) already exists" -ForegroundColor DarkGray
        continue
    }

    Write-Host "  [add ] $($col.InternalName) ($($col.Type))" -ForegroundColor Green

    $params = @{
        List         = $ListName
        InternalName = $col.InternalName
        DisplayName  = $col.DisplayName
        Type         = $col.Type
        AddToDefaultView = $false
    }
    if ($col.Choices) { $params.Choices = $col.Choices }

    Add-PnPField @params | Out-Null

    # Post-creation property tweaks
    $postValues = @{}
    if ($col.Required) { $postValues.Required = $true }
    if ($null -ne $col.DisplayFormat) { $postValues.DisplayFormat = $col.DisplayFormat }
    if ($col.Type -eq 'Note' -and $null -ne $col.PlainText) {
        $postValues.RichText = (-not $col.PlainText)
        $postValues.AppendOnly = $false
    }

    if ($postValues.Count -gt 0) {
        Set-PnPField -List $ListName -Identity $col.InternalName -Values $postValues | Out-Null
    }
}

# --- 7. Indexes -----------------------------------------------------------

$indexedColumns = @('FlowRunID', 'TargetUserUPN', 'ActionType', 'Status', 'StartTimestamp')

Write-Host ""
Write-Host "Setting indexes..." -ForegroundColor Cyan
foreach ($name in $indexedColumns) {
    Set-PnPField -List $ListName -Identity $name -Values @{ Indexed = $true } | Out-Null
    Write-Host "  [idx ] $name" -ForegroundColor Green
}

# --- 8. Views -------------------------------------------------------------

Write-Host ""
Write-Host "Creating views..." -ForegroundColor Cyan

# Helper that creates a view, or replaces it if one with the same name already exists.
function Set-AuditView {
    param(
        [string]$Title,
        [string[]]$Fields,
        [string]$Query
    )
    $existing = Get-PnPView -List $ListName -Identity $Title -ErrorAction SilentlyContinue
    if ($existing) {
        Remove-PnPView -List $ListName -Identity $Title -Force | Out-Null
    }
    Add-PnPView -List $ListName -Title $Title -Fields $Fields -Query $Query -SetAsDefault:$false | Out-Null
    Write-Host "  [view] $Title" -ForegroundColor Green
}

# Default sort used in most views: StartTimestamp descending
$sortDesc = '<OrderBy><FieldRef Name="StartTimestamp" Ascending="FALSE" /></OrderBy>'

# All Items (set as default)
$allItems = Get-PnPView -List $ListName -Identity 'All Items' -ErrorAction SilentlyContinue
if ($allItems) {
    Set-PnPView -List $ListName -Identity 'All Items' -Fields @(
        'Title','ActionType','TargetUserUPN','Status','StartTimestamp','SubmitterEmail','FlowRunURL'
    ) -Values @{ ViewQuery = $sortDesc } | Out-Null
    Write-Host "  [view] All Items (updated default)" -ForegroundColor Green
}

Set-AuditView -Title 'In Progress' -Fields @(
    'Title','ActionType','TargetUserUPN','StartTimestamp','FlowRunURL'
) -Query "<Where><Eq><FieldRef Name='Status' /><Value Type='Text'>In Progress</Value></Eq></Where>$sortDesc"

Set-AuditView -Title 'Failed' -Fields @(
    'Title','ActionType','TargetUserUPN','Status','StepsFailed','ErrorDetails','FlowRunURL','RetryCount'
) -Query "<Where><Or><Eq><FieldRef Name='Status' /><Value Type='Text'>Failed</Value></Eq><Eq><FieldRef Name='Status' /><Value Type='Text'>Partial Success</Value></Eq></Or></Where>$sortDesc"

Set-AuditView -Title 'New Hires (Last 30 Days)' -Fields @(
    'Title','TargetUserUPN','TargetUserDisplayName','Department','ManagerUPN','LicenseSKU','Status','StartDate'
) -Query "<Where><And><Eq><FieldRef Name='ActionType' /><Value Type='Text'>New Hire</Value></Eq><Geq><FieldRef Name='StartTimestamp' /><Value Type='DateTime'><Today OffsetDays='-30' /></Value></Geq></And></Where>$sortDesc"

Set-AuditView -Title 'Terminations (Last 30 Days)' -Fields @(
    'Title','TargetUserUPN','TerminationDate','MailboxConverted','MailboxForwardingTo','TerminationReason','Status'
) -Query "<Where><And><Eq><FieldRef Name='ActionType' /><Value Type='Text'>Termination</Value></Eq><Geq><FieldRef Name='StartTimestamp' /><Value Type='DateTime'><Today OffsetDays='-30' /></Value></Geq></And></Where>$sortDesc"

Set-AuditView -Title 'By Submitter' -Fields @(
    'Title','ActionType','TargetUserUPN','Status','StartTimestamp','SubmitterEmail'
) -Query '<OrderBy><FieldRef Name="SubmitterEmail" Ascending="TRUE" /><FieldRef Name="StartTimestamp" Ascending="FALSE" /></OrderBy>'

# Audit Detail view: every column. Build the field list from the schema array plus the default ones.
$allFieldNames = @('Title') + ($columns | ForEach-Object { $_.InternalName })
Set-AuditView -Title 'Audit Detail' -Fields $allFieldNames -Query $sortDesc

# --- 9. Summary -----------------------------------------------------------

Write-Host ""
Write-Host "=== Setup complete ===" -ForegroundColor Green
Write-Host ""
Write-Host "Site URL : $SiteUrl"
Write-Host "List URL : $SiteUrl/Lists/$ListName"
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Open the list, drop into the All Items view, confirm columns render."
Write-Host "  2. Add a test row manually, verify required fields are enforced, then delete it."
Write-Host "  3. Record the site URL and list URL in the README 'Tenant-specific values' table."
Write-Host "  4. Move on to Task 2: build the Microsoft Form."
