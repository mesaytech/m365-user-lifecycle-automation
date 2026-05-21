# Deliverable 2: Power Automate Flow Design

This document defines a single Power Automate flow triggered by the Microsoft Form from Deliverable 1. The flow branches on `Action Type`, calls Microsoft Graph for all identity operations, writes to the `Lifecycle Audit Log` SharePoint list at every step, and runs parallel error-handler branches with retry on HTTP 429.

The flow is named `Flow-M365-Lifecycle-Main`.

---

## 1. Architecture overview

```
+------------------------------+
| Microsoft Forms trigger      |
| "When a new response..."     |
+--------------+---------------+
               |
+--------------v---------------+
| Get response details         |
+--------------+---------------+
               |
+--------------v---------------+
| Initialize variables         |
| Create initial audit row     |
+--------------+---------------+
               |
+--------------v---------------+
| Switch on Action Type        |
+----+--------------+----------+
     |              |
+----v-----+   +----v-----+
| New Hire |   | Termin.  |
| Scope    |   | Scope    |
+----+-----+   +----+-----+
     |              |
     +------+-------+
            |
+-----------v------------+
| Finalize Scope         |
| (runs after either)    |
+-----------+------------+
            |
+-----------v------------+
| Error Handler Scope    |
| (runs ONLY if any of   |
|  the above failed)     |
+------------------------+
```

The Finalize scope and Error Handler scope both update the audit row. They differ only in which fields they set and in their "configure run after" conditions.

---

## 2. Prerequisites

### 2.1 Entra ID app registration

Create an app registration that the flow uses to call Microsoft Graph. The Office 365 Outlook connector uses a separate delegated connection (your admin account) and does not need this.

This app registration is dedicated to the Power Automate flow's Graph calls. It is **not** the same app as the one used by `scripts/Setup-SharePointList.ps1`. PnP.PowerShell 2.x requires its own app registration (the bundled Management Shell app was removed and the `PNPPOWERSHELL_CLIENTID` environment variable is not honored in current builds), and the setup script takes its ClientId via the `-ClientId` parameter. See the README's "SharePoint" prerequisites section for the `Register-PnPEntraIDApp` one-liner that creates that second app. Keep the two ClientIds straight, the flow's app needs application Graph permissions while the PnP app needs delegated SharePoint permissions.

In the Entra admin center at https://entra.microsoft.com:

1. Identity → Applications → App registrations → New registration
2. Name: `M365-Lifecycle-Automation`
3. Supported account types: "Accounts in this organizational directory only"
4. Redirect URI: leave blank
5. Click Register

Record these three values from the Overview page, you will paste them into Power Automate connections:

- Application (client) ID
- Directory (tenant) ID
- Tenant primary domain (used to construct UPNs, format `{tenantname}.onmicrosoft.com` in the dev tenant)

Create a client secret:

1. Certificates and secrets → New client secret
2. Description: `PowerAutomate-Lifecycle`
3. Expires: 6 months (rotate before expiry, set a calendar reminder)
4. Copy the secret value immediately. It is shown only once.

### 2.2 API permissions

Add these as Application permissions (not Delegated) on the app registration, then click "Grant admin consent for {tenant}".

| API | Permission | Used for |
|---|---|---|
| Microsoft Graph | `User.ReadWrite.All` | Create users, update users, block sign-in, revoke sessions |
| Microsoft Graph | `Directory.ReadWrite.All` | Assign and remove licenses, read tenant `subscribedSkus` |
| Microsoft Graph | `Group.ReadWrite.All` | Resolve groups by display name, add and remove members |
| Microsoft Graph | `Organization.Read.All` | Read `subscribedSkus` (covered by Directory.ReadWrite.All but listed for clarity if you ever tighten permissions) |
| Microsoft Graph | `MailboxSettings.ReadWrite` | Set automatic forwarding via `mailboxSettings` |
| Microsoft Graph | `Mail.ReadWrite` | Create the inbox forwarding rule on the terminated user's mailbox |

For the shared mailbox conversion (Exchange Online cmdlet `Set-Mailbox -Type Shared`), Graph has no equivalent endpoint. That step uses an Azure Automation runbook described in section 2.4.

### 2.3 Power Automate connections

In Power Automate at https://make.powerautomate.com, create these connections under My Connections before building the flow:

| Connection | Purpose | How to authenticate |
|---|---|---|
| Microsoft Forms | Trigger | Your admin account |
| SharePoint | Audit log read/write | Your admin account |
| Office 365 Outlook | Send welcome and notification emails | A shared mailbox the user has Send As on, or your admin account (mention the actual From address in section 6.10) |
| HTTP | Graph calls. The HTTP action uses inline OAuth, no pre-built connection. | n/a (configured per action) |
| Azure Automation | Shared mailbox conversion | Service principal of the same app registration, granted Contributor on the automation account |

### 2.4 Azure Automation account (only needed if any termination will choose "Convert to shared mailbox = Yes")

Power Automate cannot run `Set-Mailbox` directly. The standard pattern is an Azure Automation runbook with the ExchangeOnlineManagement module and certificate-based auth.

If your dev tenant has no Azure subscription, sign up for the free Azure trial linked to the same admin account. Azure Automation has a free tier of 500 job-execution minutes per month, which is more than enough.

Setup steps:

1. Create a resource group `rg-m365-automation` in your nearest region.
2. Create an Automation account `aa-m365-lifecycle`.
3. Identity → System-assigned managed identity → On. Save the object (principal) ID, you will use it in steps 5 and 6.
4. Modules → import `ExchangeOnlineManagement` **version 3.4.0** to runtime version **7.2**. Do not use the latest version, see "Module version pinning" below for the reason.
5. Grant the managed identity the **Exchange Recipient Administrator** directory role in Entra. From Entra admin center → Identity → Roles & admins → All roles → search "Exchange Recipient Administrator" → click the role → Add assignments → search the managed identity by its object ID from step 3 → Add.
6. Grant the managed identity the **Exchange.ManageAsApp** application permission on the Office 365 Exchange Online API. The portal does not surface this for managed identities so this is a PowerShell-only step, see "Granting Exchange.ManageAsApp" below.
7. Runbooks → Create a runbook → name `Convert-MailboxToShared` → type PowerShell, runtime version 7.2.
8. Paste the runbook code below into the editor, Save, then Publish.

**Module version pinning**

ExchangeOnlineManagement 3.6.0 and later introduced .NET 8 dependencies. The Azure Automation PS 7.2 sandbox runs on .NET 6, so newer versions of the module return `HRESULT 0x80131047` (`FUSION_E_INVALID_NAME`) when the runbook tries to call any cmdlet. Version 3.4.0 is the last release that targets only .NET 6 and loads cleanly.

The portal's "Browse from gallery" experience always installs the latest version of a module and gives no version selector, so you cannot pin a specific version through the UI. Download the version locally and upload the zip instead:

```powershell
$tempPath = "C:\temp\modules"
New-Item -ItemType Directory -Path $tempPath -Force | Out-Null
Save-Module -Name ExchangeOnlineManagement -RequiredVersion 3.4.0 -Path $tempPath -Repository PSGallery

$sourceDir = "$tempPath\ExchangeOnlineManagement\3.4.0"
$destZip = "$tempPath\ExchangeOnlineManagement.zip"
Compress-Archive -Path "$sourceDir\*" -DestinationPath $destZip -Force
```

Then in the Automation account: Modules → + Add a module → Upload a module file → select `ExchangeOnlineManagement.zip` → Runtime version 7.2 → Import. Wait for status to flip to `Available`.

**Granting Exchange.ManageAsApp**

`Connect-ExchangeOnline -ManagedIdentity` returns `UnAuthorized (UnAuthorized)` if the managed identity lacks the `Exchange.ManageAsApp` application role on the Office 365 Exchange Online API service principal. The Entra directory role from step 5 covers recipient operations but does not grant the Exchange Online API access by itself.

Run this once after the managed identity exists:

```powershell
Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All", "Directory.Read.All" -NoWelcome

# Managed identity SP
$miSp = Get-MgServicePrincipal -Filter "displayName eq 'aa-m365-lifecycle'"
if (-not $miSp) { throw "Managed identity SP not found" }

# Office 365 Exchange Online API SP (well-known appId)
$exoSp = Get-MgServicePrincipal -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'"

# Exchange.ManageAsApp role ID (well-known)
$appRoleId = "dc50a0fb-09a3-484d-be87-e023b12c6440"

New-MgServicePrincipalAppRoleAssignment `
  -ServicePrincipalId $miSp.Id `
  -PrincipalId $miSp.Id `
  -ResourceId $exoSp.Id `
  -AppRoleId $appRoleId | Out-Null

Disconnect-MgGraph | Out-Null
```

The PowerShell session must sign in as a Global Admin (the `AppRoleAssignment.ReadWrite.All` scope requires admin consent at sign-in). Wait ~5 minutes after the grant for propagation before testing the runbook.

```powershell
param(
    [Parameter(Mandatory=$true)][string]$UserPrincipalName
)

$ErrorActionPreference = 'Stop'

# Verified tenant domain. Connect-ExchangeOnline accepts either the tenant GUID
# or the verified primary domain for -Organization. Using the domain avoids a
# dependency on Az.Accounts and Connect-AzAccount inside the runbook.
$tenantDomain = 'cloudopslabs.onmicrosoft.com'

Connect-ExchangeOnline -ManagedIdentity -Organization $tenantDomain

try {
    Set-Mailbox -Identity $UserPrincipalName -Type Shared -ErrorAction Stop
    Write-Output "Converted $UserPrincipalName to shared mailbox."
}
catch {
    # Note ${UserPrincipalName} — PowerShell parses $UserPrincipalName: as a scoped
    # variable reference and rejects the string. The braces delimit the name.
    Write-Error "Failed to convert ${UserPrincipalName}: $($_.Exception.Message)"
    throw
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
}
```

If this runbook is reused in a different tenant, change `$tenantDomain` to that tenant's verified primary domain, or parameterize it via an additional `param()` entry.

The flow invokes this runbook in step 7.9.

---

## 3. Trigger and initialization

### 3.1 Trigger: When a new response is submitted

- Connector: Microsoft Forms
- Form ID: select the form built in Deliverable 1
- Concurrency control: default (no degree of parallelism set)

The trigger only fires once per submission and returns a Response ID. The Forms trigger does not return the answer values, that is the next step.

### 3.2 Get response details

- Connector: Microsoft Forms
- Action: Get response details
- Form ID: same as trigger
- Response ID: dynamic content `Response Id` from trigger

This is the action that returns every answer keyed by question identifier. Use the "Add dynamic content" picker rather than hand-typing the keys, the GUIDs differ per form.

### 3.3 Initialize variables

Add nine Initialize variable actions in sequence at the top of the flow. They must be at the top level (not inside a scope) because Power Automate variables are flow-global and can only be initialized at the root.

| Variable name | Type | Initial value (expression) |
|---|---|---|
| `varAuditRecordID` | String | `concat(if(equals(outputs('Get_response_details')?['body/{question-id-Q1}'], 'New Hire'), 'NH-', 'TERM-'), formatDateTime(utcNow(), 'yyyyMMdd-HHmmss'), '-', substring(workflow()['run']['name'], 0, 8))` |
| `varStartTimeUTC` | String | `utcNow()` |
| `varAuditRowID` | Integer | `0` (set after the audit row is created in 3.4) |
| `varStepsCompleted` | String | `''` |
| `varStepsFailed` | String | `''` |
| `varErrorDetails` | String | `''` |
| `varRetryCount` | Integer | `0` |
| `varTargetUserObjectID` | String | `''` |
| `varTempPassword` | String | `concat(substring(replace(guid(), '-', ''), 0, 11), 'A', '7', '!')` |

The `varTempPassword` expression produces a 14-character password. The 11-character GUID slice (with hyphens stripped) provides lowercase letters a-f and digits. The literal `A`, `7`, and `!` appended at the end guarantee at least one uppercase letter, one digit, and one symbol, regardless of what the GUID slice happens to contain. This satisfies the Entra ID default password policy requirement of 3 of 4 character classes (it actually delivers all 4). If your tenant has a custom password policy with longer minimums or banned-substring rules, lengthen the GUID slice and review the appended characters.

Replace `{question-id-Q1}` with the actual Forms question ID. You get the ID by opening the Get response details action's dynamic content list, hovering over "Action type", and copying the technical name.

### 3.4 Create initial audit row

- Connector: SharePoint
- Action: Create item
- Site Address: your audit site URL
- List Name: `LifecycleAuditLog`
- Field values:

| List column | Value (dynamic content or expression) |
|---|---|
| Title | `variables('varAuditRecordID')` |
| ActionType Value | `Action type` (from Get response details) |
| TargetUserUPN | For New Hire: constructed in step 6.1 then patched in via Update item. For Termination: `User to offboard (UPN)` from Q13. Leave empty here and set it after the branch knows the UPN. Actually set it now using a coalesce expression: `coalesce(outputs('Get_response_details')?['body/{question-id-Q13}'], '')` and update later for new hires. |
| SubmitterEmail | `Responder's Email` |
| SubmitterDisplayName | `Name` |
| FormResponseID | `Response Id` (from trigger) |
| FlowRunID | `workflow()['run']['name']` |
| FlowRunURL | `concat('https://make.powerautomate.com/manage/environments/', workflow()['tags']['environmentName'], '/flows/', workflow()['name'], '/runs/', workflow()['run']['name'])` |
| StartTimestamp | `variables('varStartTimeUTC')` |
| Status Value | `In Progress` |

After this action, set `varAuditRowID` to the returned `ID` of the created item using a Set variable action. Every subsequent audit update references this row ID.

### 3.5 Switch on Action Type

- Action: Switch
- On: `Action type` (from Get response details)
- Case 1 value: `New Hire` → contains the New Hire scope (section 6)
- Case 2 value: `Termination` → contains the Termination scope (section 7)
- Default case: a single "Terminate" action with status `Failed` and message `Unknown action type received from form: @{outputs('Get_response_details')?['body/{question-id-Q1}']}`

---

## 4. Standard step pattern (read this before sections 6 and 7)

Every Graph call in the New Hire and Termination scopes uses the same five-block pattern:

```
+------------------------------------+
| (1) HTTP action - call Graph       |
|     with retry policy              |
+----+----------------+--------------+
     | success        | fail
+----v---------+  +---v-------------+
| (2) Update   |  | (3) Append to   |
| audit row    |  | varStepsFailed  |
| StepsCompl.  |  | and varErrorDet.|
+--------------+  +---+-------------+
                      |
                  +---v-------------+
                  | (4) Update audit|
                  | row with        |
                  | partial state   |
                  +---+-------------+
                      |
                  +---v-------------+
                  | (5) Continue or |
                  | Terminate based |
                  | on step criti-  |
                  | cality          |
                  +-----------------+
```

The "configure run after" settings:

- Block 2 (success append) runs only after block 1 has "is successful"
- Block 3 (failure append) runs after block 1 has "has failed", "has timed out", or "is skipped"
- Block 4 runs after block 3 has "is successful"
- Block 5 is a Terminate action with status `Failed` for critical steps, or it is omitted for non-critical steps that should not stop the flow

### 4.1 HTTP action template (Graph)

Every Graph call uses the HTTP premium action configured this way:

| Field | Value |
|---|---|
| Method | (varies, see each step) |
| URI | (varies, see each step) |
| Headers | `Content-Type: application/json` |
| Authentication | Active Directory OAuth |
| Authority | `https://login.microsoftonline.com` |
| Tenant | your tenant ID from 2.1 |
| Audience | `https://graph.microsoft.com` |
| Client ID | your app ID from 2.1 |
| Credential Type | Secret |
| Secret | your client secret from 2.1 (paste into the connection, do not hard-code in the action) |

In the action's "Settings" pane, set Retry Policy:

- Type: Exponential interval
- Count: 4
- Interval: PT10S
- Minimum interval: PT5S
- Maximum interval: PT1H

This policy retries on any 408, 429, 5xx response. Graph's `Retry-After` header is not honored by the built-in policy. For the 429-specific Retry-After case described in section 8, wrap the action in a Do until loop.

### 4.2 Critical vs non-critical steps

| Step type | Failure behavior | Examples |
|---|---|---|
| Critical | Append to StepsFailed, set audit Status to `Failed`, terminate the flow run | Create user, look up termination target user, block sign-in |
| Recoverable | Append to StepsFailed, set audit Status to `Partial Success`, continue the flow | Add a single group member, send welcome email |

Section 6 and section 7 mark each step C (critical) or R (recoverable).

---

## 5. Helper expression reference

| Need | Expression |
|---|---|
| Append step to varStepsCompleted | Set variable `varStepsCompleted` to `concat(variables('varStepsCompleted'), 'StepName: success', char(10))` |
| Append step to varStepsFailed | Set variable `varStepsFailed` to `concat(variables('varStepsFailed'), 'StepName: failed', char(10))` |
| Append error body | Set variable `varErrorDetails` to `concat(variables('varErrorDetails'), 'StepName error: ', string(body('HTTP_Action_Name')), char(10))` |
| Increment retry counter | Set variable `varRetryCount` to `add(variables('varRetryCount'), 1)` |
| Compose UPN | `concat(toLower(replace(outputs('Get_response_details')?['body/{q-firstname}'], ' ', '.')), '.', toLower(replace(outputs('Get_response_details')?['body/{q-lastname}'], ' ', '.')), '@', 'yourtenant.onmicrosoft.com')` |
| Compose mailNickname | `concat(toLower(outputs('Get_response_details')?['body/{q-firstname}']), '.', toLower(outputs('Get_response_details')?['body/{q-lastname}']))` |
| Compose display name | `if(empty(outputs('Get_response_details')?['body/{q-preferreddisplay}']), concat(outputs('Get_response_details')?['body/{q-firstname}'], ' ', outputs('Get_response_details')?['body/{q-lastname}']), outputs('Get_response_details')?['body/{q-preferreddisplay}'])` |
| Combine termination date and time | `formatDateTime(addHours(startOfDay(outputs('Get_response_details')?['body/{q-termdate}']), if(equals(outputs('Get_response_details')?['body/{q-termtime}'], 'Immediate'), 0, int(substring(outputs('Get_response_details')?['body/{q-termtime}'], 0, 2)))), 'yyyy-MM-ddTHH:mm:ssZ')` |

---

## 6. New Hire branch

Scope name: `Scope_NewHire`. Every action below sits inside this scope unless noted.

### 6.1 Compose UPN, mailNickname, display name

Three Compose actions named `Compose_UPN`, `Compose_MailNickname`, `Compose_DisplayName`, using the expressions in section 5.

Critical: C

### 6.2 Create user

| Field | Value |
|---|---|
| Action name | `HTTP_CreateUser` |
| Method | POST |
| URI | `https://graph.microsoft.com/v1.0/users` |
| Body | (below) |
| Critical | C |

```json
{
  "accountEnabled": true,
  "displayName": "@{outputs('Compose_DisplayName')}",
  "mailNickname": "@{outputs('Compose_MailNickname')}",
  "userPrincipalName": "@{outputs('Compose_UPN')}",
  "passwordProfile": {
    "forceChangePasswordNextSignIn": true,
    "password": "@{variables('varTempPassword')}"
  },
  "jobTitle": "@{outputs('Get_response_details')?['body/{q-jobtitle}']}",
  "department": "@{outputs('Get_response_details')?['body/{q-department}']}",
  "officeLocation": "@{outputs('Get_response_details')?['body/{q-office}']}",
  "mobilePhone": "@{outputs('Get_response_details')?['body/{q-mobile}']}",
  "usageLocation": "US"
}
```

`usageLocation` is required before assignLicense will accept the call. Set it to your tenant's primary country, two-letter ISO code.

After success: Set `varTargetUserObjectID` to `body('HTTP_CreateUser')?['id']`. Update audit row to set `TargetUserObjectID`, `TargetUserUPN` (from `outputs('Compose_UPN')`), `TargetUserDisplayName`.

### 6.3 Assign manager

| Field | Value |
|---|---|
| Action name | `HTTP_GetManager` |
| Method | GET |
| URI | `https://graph.microsoft.com/v1.0/users/@{outputs('Get_response_details')?['body/{q-managerupn}']}?$select=id` |
| Critical | R |

Then `HTTP_SetManager`:

| Field | Value |
|---|---|
| Method | PUT |
| URI | `https://graph.microsoft.com/v1.0/users/@{variables('varTargetUserObjectID')}/manager/$ref` |
| Body | (below) |

```json
{
  "@odata.id": "https://graph.microsoft.com/v1.0/users/@{body('HTTP_GetManager')?['id']}"
}
```

### 6.4 Look up license SKU ID

| Field | Value |
|---|---|
| Action name | `HTTP_GetSkus` |
| Method | GET |
| URI | `https://graph.microsoft.com/v1.0/subscribedSkus?$select=skuId,skuPartNumber` |
| Critical | C |

Then a Filter array action `Filter_TargetSku`:

- From: `body('HTTP_GetSkus')?['value']`
- Condition: `item()?['skuPartNumber']` is equal to `outputs('Get_response_details')?['body/{q-license}']`

Then Compose `Compose_SkuId`: `first(body('Filter_TargetSku'))?['skuId']`. Add a condition that terminates with `Failed` if the filter returns empty (license not found in tenant).

### 6.5 Assign license

| Field | Value |
|---|---|
| Action name | `HTTP_AssignLicense` |
| Method | POST |
| URI | `https://graph.microsoft.com/v1.0/users/@{variables('varTargetUserObjectID')}/assignLicense` |
| Critical | C |

```json
{
  "addLicenses": [
    { "skuId": "@{outputs('Compose_SkuId')}", "disabledPlans": [] }
  ],
  "removeLicenses": []
}
```

### 6.6 Add to groups

The Forms Q11 returns a single newline-delimited string. Split it and loop.

1. Compose `Compose_GroupList`: `split(outputs('Get_response_details')?['body/{q-groups}'], decodeUriComponent('%0A'))`
2. Apply to each group name (with concurrency control disabled, set to 1, so failures are easier to diagnose):
   - `HTTP_GetGroup`: GET `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '@{replace(trim(item()), '''', '''''')}'&$select=id` (R). The `replace` doubles any single-apostrophe in the group name to escape it for OData. In Power Automate's expression language a single quote inside a string is written as two single quotes, so the literal `'` is `''''` and the literal `''` is `''''''` in the expression. The example `O'Brien Team` becomes the OData filter value `'O''Brien Team'`, which Graph parses correctly.
   - Condition: `length(body('HTTP_GetGroup')?['value'])` greater than 0
     - If yes: `HTTP_AddGroupMember`: POST `https://graph.microsoft.com/v1.0/groups/@{first(body('HTTP_GetGroup')?['value'])?['id']}/members/$ref` (R) with body:
     - If no: append to varStepsFailed with `Group not found: @{item()}`
3. After each successful add, append the group name to a `varGroupsAdded` accumulator variable

```json
{
  "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/@{variables('varTargetUserObjectID')}"
}
```

### 6.7 Send welcome email to new hire

- Connector: Office 365 Outlook
- Action: Send an email (V2)
- To: `outputs('Compose_UPN')`
- Cc: `outputs('Get_response_details')?['body/{q-managerupn}']`
- Subject: `Welcome to the team, @{outputs('Get_response_details')?['body/{q-firstname}']}`
- Body: HTML, include the temp password from `variables('varTempPassword')`, sign-in URL `https://portal.office.com`, and the start date from Q9. Mention they will be required to change the password on first sign-in.
- Critical: R

The new user's mailbox is provisioned asynchronously by Exchange Online after license assignment and may not exist yet when this email is sent. Send via your admin From address (the connection's user), not via the new user's mailbox.

### 6.8 New Hire scope complete

After 6.7, append `NewHire branch complete` to `varStepsCompleted`. The Finalize scope (section 9) will set Status and EndTimestamp.

---

## 7. Termination branch

Scope name: `Scope_Termination`.

### 7.1 Resolve target user

| Field | Value |
|---|---|
| Action name | `HTTP_GetTargetUser` |
| Method | GET |
| URI | `https://graph.microsoft.com/v1.0/users/@{outputs('Get_response_details')?['body/{q-termupn}']}?$select=id,displayName,assignedLicenses,userPrincipalName` |
| Critical | C |

After success: Set `varTargetUserObjectID` to `body('HTTP_GetTargetUser')?['id']`. Update audit row with `TargetUserObjectID` and `TargetUserDisplayName`.

### 7.2 Delay until scheduled termination time (if not Immediate)

- Action: Condition `outputs('Get_response_details')?['body/{q-termtime}']` is equal to `Immediate`
- If yes branch: do nothing, continue
- If no branch: Delay until action with `Timestamp` set to the combined date-time expression from section 5

### 7.3 Block sign-in

| Field | Value |
|---|---|
| Action name | `HTTP_BlockSignIn` |
| Method | PATCH |
| URI | `https://graph.microsoft.com/v1.0/users/@{variables('varTargetUserObjectID')}` |
| Critical | C |

```json
{ "accountEnabled": false }
```

### 7.4 Revoke all active sessions

| Field | Value |
|---|---|
| Action name | `HTTP_RevokeSessions` |
| Method | POST |
| URI | `https://graph.microsoft.com/v1.0/users/@{variables('varTargetUserObjectID')}/revokeSignInSessions` |
| Body | (none) |
| Critical | C |

### 7.5 Remove from all groups

The `memberOf` endpoint paginates at 100 entries per page. A real-tenant user can easily exceed this (one M365 group per project, plus security groups for shared mailboxes, distribution lists, and Teams). The flow follows `@odata.nextLink` until null so termination is complete regardless of count.

Step structure:

1. Initialize two top-level variables (add to the section 3.3 list):
   - `varNextLink` (String) with initial value `concat('https://graph.microsoft.com/v1.0/users/', variables('varTargetUserObjectID'), '/memberOf?$select=id,displayName,@odata.type&$top=100')`
   - `varAllMemberships` (Array) with initial value `[]`

2. Do until loop `DoUntil_Memberships`:
   - Condition: `empty(variables('varNextLink'))` is equal to `true`
   - Limit: Count 100, Timeout PT1H

3. Inside the Do until loop:
   - `HTTP_GetMembershipsPage` (C): GET `@{variables('varNextLink')}`. Use the Authentication block from section 4.1.
   - Set variable `varAllMemberships` to `union(variables('varAllMemberships'), body('HTTP_GetMembershipsPage')?['value'])`. The `union` function de-duplicates if a group ever appears twice across pages (it should not, but this is defensive).
   - Set variable `varNextLink` to `coalesce(body('HTTP_GetMembershipsPage')?['@odata.nextLink'], '')`. When the response omits `@odata.nextLink`, `coalesce` returns the empty string, which ends the loop.

4. Apply to each membership in `variables('varAllMemberships')`:
   - Condition: `item()?['@odata.type']` is equal to `#microsoft.graph.group` (skip directory roles, which appear here too and use a different removal endpoint)
   - If yes: `HTTP_RemoveMember` (R): DELETE `https://graph.microsoft.com/v1.0/groups/@{item()?['id']}/members/@{variables('varTargetUserObjectID')}/$ref`
   - Append the group display name to `varGroupsRemoved` accumulator on success

Dynamic distribution groups, on-premises synced groups, and the special "All Users" group will return errors here. These are expected, log them to varStepsFailed but do not terminate.

The Do until count limit of 100 covers users with up to 10,000 group memberships (100 pages × 100 entries), well beyond any realistic case. If the loop ever exhausts its limit without `varNextLink` becoming empty, the flow continues and removes whatever it gathered, then logs `Pagination limit reached for memberOf` to varStepsFailed.

### 7.6 Remove all licenses

1. Compose `Compose_LicenseIdsToRemove`: `xpath(xml(json(concat('{"licenses":', string(body('HTTP_GetTargetUser')?['assignedLicenses']), '}'))), '/licenses/skuId/text()')`

   Simpler equivalent using Select action: From `body('HTTP_GetTargetUser')?['assignedLicenses']`, Map `item()?['skuId']`. The output is the array needed below.

2. `HTTP_RemoveLicenses`: POST `https://graph.microsoft.com/v1.0/users/@{variables('varTargetUserObjectID')}/assignLicense` (C)

```json
{
  "addLicenses": [],
  "removeLicenses": @{body('Select_LicenseIds')}
}
```

### 7.7 Set mail forwarding (if Q17 = Yes)

Condition: `outputs('Get_response_details')?['body/{q-forward}']` is equal to `Yes`

If yes:

| Field | Value |
|---|---|
| Action name | `HTTP_SetForwarding` |
| Method | PATCH |
| URI | `https://graph.microsoft.com/v1.0/users/@{variables('varTargetUserObjectID')}/mailboxSettings` |
| Critical | R |

```json
{
  "automaticRepliesSetting": {
    "status": "alwaysEnabled",
    "internalReplyMessage": "I am no longer with the company. For assistance, contact @{outputs('Get_response_details')?['body/{q-delegate}']}.",
    "externalReplyMessage": "I am no longer with the company. For assistance, contact @{outputs('Get_response_details')?['body/{q-delegate}']}.",
    "externalAudience": "all"
  }
}
```

Then create an inbox rule to actually forward incoming mail:

| Field | Value |
|---|---|
| Action name | `HTTP_CreateForwardRule` |
| Method | POST |
| URI | `https://graph.microsoft.com/v1.0/users/@{variables('varTargetUserObjectID')}/mailFolders/inbox/messageRules` |
| Critical | R |

```json
{
  "displayName": "Forward to delegate after termination",
  "sequence": 1,
  "isEnabled": true,
  "actions": {
    "forwardTo": [
      {
        "emailAddress": {
          "address": "@{outputs('Get_response_details')?['body/{q-delegate}']}"
        }
      }
    ],
    "stopProcessingRules": false
  }
}
```

### 7.8 Convert mailbox to shared (if Q16 = Yes)

Condition: `outputs('Get_response_details')?['body/{q-sharedmbx}']` is equal to `Yes`

If yes:

- Action: Azure Automation - Create job
- Subscription, Resource group, Automation account: from section 2.4
- Runbook name: `Convert-MailboxToShared`
- Wait for job: Yes
- Runbook parameters: `UserPrincipalName` = `outputs('Get_response_details')?['body/{q-termupn}']`
- Critical: R

After the job, query the job output by adding a Get job output action and check for the success string. If the output starts with `Failed`, append to varStepsFailed.

Update audit row to set `MailboxConverted` = Yes if the job succeeded.

### 7.9 Send notification email

- Connector: Office 365 Outlook → Send an email (V2)
- To: a static IT distribution list, configurable
- Cc: the manager UPN if known (look up via `HTTP_GetTargetUser` response if you added `manager` to `$expand` in 7.1)
- Subject: `Offboarding complete: @{body('HTTP_GetTargetUser')?['displayName']}`
- Body: include UPN, termination date, list of groups removed (from `varGroupsRemoved`), mailbox conversion status, forwarding status

Critical: R

### 7.10 Termination scope complete

Append `Termination branch complete` to `varStepsCompleted`.

---

## 8. Special case: Retry-After on HTTP 429

Graph's documented throttling pattern is to return 429 with a `Retry-After` header in seconds. The built-in HTTP retry policy uses exponential backoff and ignores this header. In most dev tenant scenarios the built-in policy is enough.

For the user-creation step (6.2) and the bulk group operations (6.6, 7.5), where 429 is most likely, wrap the action in a Do until loop:

1. Initialize a boolean variable `varRequestSucceeded` to `false` and an integer `varAttempt` to `0` at the top of the loop
2. Do until: `varRequestSucceeded` is true OR `varAttempt` is greater than 5
3. Inside:
   - HTTP action (turn off the action's built-in retry policy to avoid double-retry)
   - Condition: `outputs('HTTP_Action')?['statusCode']` is equal to 429
     - If yes: Delay action, Count = `int(outputs('HTTP_Action')?['headers']?['Retry-After'])`, Unit = Second. Then increment `varAttempt` and `varRetryCount`.
     - If no: Set `varRequestSucceeded` to true

The `varRetryCount` value is written to the audit row in the Finalize scope so you can see throttling patterns.

---

## 9. Finalize scope and error handler

### 9.1 Finalize scope

- Scope name: `Scope_Finalize`
- Configure run after: Switch (3.5) has succeeded, has failed, or has timed out

Inside the scope, a single SharePoint Update item action:

| Column | Value |
|---|---|
| Id | `variables('varAuditRowID')` |
| EndTimestamp | `utcNow()` |
| DurationSeconds | `div(sub(ticks(utcNow()), ticks(variables('varStartTimeUTC'))), 10000000)` |
| Status Value | Determined by a Compose: if `varStepsFailed` is empty then `Succeeded`, else if any critical step failed then `Failed`, else `Partial Success` |
| StepsCompleted | `variables('varStepsCompleted')` |
| StepsFailed | `variables('varStepsFailed')` |
| ErrorDetails | `variables('varErrorDetails')` |
| RetryCount | `variables('varRetryCount')` |
| GroupsAdded | `variables('varGroupsAdded')` (only meaningful for new hires, leave for terminations) |
| GroupsRemoved | `variables('varGroupsRemoved')` (only meaningful for terminations) |
| TargetUserObjectID | `variables('varTargetUserObjectID')` |

The expression for Status:

```
if(
  and(equals(variables('varStepsFailed'), ''), equals(variables('varErrorDetails'), '')),
  'Succeeded',
  if(
    contains(variables('varStepsFailed'), 'CRITICAL'),
    'Failed',
    'Partial Success'
  )
)
```

Tag critical failures with the literal string `CRITICAL` when you append to `varStepsFailed` in the critical-step error branches, so the expression above can distinguish them.

### 9.2 Top-level error handler scope

- Scope name: `Scope_ErrorHandler`
- Configure run after: `Scope_Finalize` has failed OR has timed out OR is skipped (covers the case where the trigger or initialization itself failed and the finalize step never ran)

Inside: a single Send an email (V2) to your admin email, subject `Lifecycle automation failure: @{variables('varAuditRecordID')}`, body containing the flow run URL, error details, and a link to the audit row.

This is a safety net. If reached, the audit row may not have been updated, the email is the only signal.

---

## 10. Testing plan

Test in this order. Do not move to step N+1 until step N passes.

1. Submit a New Hire form with all fields filled, including 2 groups and a phone number
   - Verify audit row created with Status = In Progress within 5 seconds
   - Verify user appears in Entra ID with correct attributes
   - Verify license appears in the user's assigned licenses
   - Verify manager is set
   - Verify both groups list the new user as a member
   - Verify welcome email arrives at the new mailbox within 10 minutes (license provisioning lag)
   - Verify audit row updates to Status = Succeeded with no StepsFailed
2. Submit a New Hire form with one group name misspelled
   - Verify Status = Partial Success
   - Verify the misspelled group is named in StepsFailed
   - Verify the user is still created and the other group is added
3. Submit a New Hire form with a manager UPN that does not exist
   - Verify Status = Partial Success
   - Verify HTTP_GetManager error captured in ErrorDetails
4. Submit a Termination form with Q15 = Immediate, Q16 = No, Q17 = No
   - Verify sign-in is blocked within 60 seconds
   - Verify all groups stripped (verify in Entra ID memberOf)
   - Verify licenses removed
5. Submit a Termination form with Q15 = 17:00, Q16 = Yes, Q17 = Yes, Q18 = your own address
   - Verify the flow waits at Delay until until the scheduled time
   - At the scheduled time, verify mailbox is shared (`Get-Mailbox -Identity x | Select RecipientTypeDetails`)
   - Verify forwarding rule exists (`Get-InboxRule -Mailbox x`)
6. Force a 429 by submitting six new-hire forms in rapid succession
   - Verify at least one run records RetryCount > 0
   - Verify all six eventually reach Succeeded or Partial Success

Document each test's audit row ID in the README.

---

## Build checklist

- [ ] Entra ID app `M365-Lifecycle-Automation` registered, client secret stored
- [ ] All 6 Graph application permissions granted, admin consent given
- [ ] Power Automate connections created for Forms, SharePoint, Outlook
- [ ] Azure Automation account created, ExchangeOnlineManagement module installed
- [ ] `Convert-MailboxToShared` runbook published
- [ ] Managed identity granted Exchange Recipient Administrator
- [ ] Flow `Flow-M365-Lifecycle-Main` created with the trigger and Get response details
- [ ] All 9 variables initialized at the top level
- [ ] Initial audit row creation step working (test by saving and triggering once)
- [ ] Switch action with both cases and a default
- [ ] New Hire scope with all 7 numbered steps (6.1 through 6.7)
- [ ] Termination scope with all 10 numbered steps (7.1 through 7.10)
- [ ] Every Graph step has a parallel error-handler branch (block 3 in the section 4 pattern)
- [ ] Retry policy set to exponential, count 4, on every HTTP action
- [ ] Do until 429-handling wrapper added to create-user, add-group-member, and remove-group-member steps
- [ ] Finalize scope updates the audit row with Status, EndTimestamp, Duration
- [ ] Top-level error handler scope sends a failure email
- [ ] All 6 test scenarios passed and documented
