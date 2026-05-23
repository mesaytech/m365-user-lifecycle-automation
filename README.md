# M365 User Lifecycle Automation

Reference implementation of an end-to-end Power Automate flow that handles new-hire provisioning and termination offboarding in a Microsoft 365 tenant. Includes full audit logging in SharePoint, parallel error-handler branches around every Microsoft Graph call, and a PowerShell-based JSON-edit tooling workflow.

Intended as a deployable reference for other tenants. All tenant-specific identifiers are externalized to `config.local.json` (gitignored). See **Prerequisites** and **Tenant-specific values to capture** sections below.

First built and validated in a Microsoft 365 Developer Program tenant (Business Premium, 25 licenses).

**Author:** Mesay Bezuneh

## Status

| Day | Milestone | State |
|---|---|---|
| 1 | Tenant activated, SharePoint and Forms schema designed | Done |
| 1 | Power Automate flow design (both branches, error handling, runbook, pagination) documented | Done |
| 2 | Graph API reference compiled | Done |
| 2 | Microsoft Form built in tenant | Done |
| 2 | SharePoint list created in tenant | Done |
| 2 | Entra ID app registration + 3 Graph Application permissions + admin consent | Done |
| 2 | Azure Automation account + runbook + role + API permission (smoke test passed) | Done |
| 3 | Flow Phase A + B (trigger, variables, audit row, Switch) | Done |
| 3 | New Hire scope §6.1–§6.5 (provision user, manager wiring, license assignment) | Done — verified end-to-end |
| 3 | New Hire scope §6.6 (group memberships) | Done — verified end-to-end (success + not-found paths) |
| 3 | New Hire scope §6.7 (welcome email) | Done — verified end-to-end |
| 3 | New Hire scope §6.8 (finalize audit row, Status=Succeeded/Partial) | Done — verified both paths |
| 4 | Termination scope §7.1 (resolve target user) | Done — verified end-to-end |
| 4 | Termination scope §7.2 (Delay Until with HH:MM / Immediate validation) | Done — verified strict invalid path |
| 4 | Termination scope §7.3 (block sign-in) | Done — verified accountEnabled=false in Entra |
| 4 | Termination scope §7.4 (revoke active sessions) | Done — verified signInSessionsValidFromDateTime updated |
| 4 | Termination scope §7.5 (remove from all groups, paginated loop) | Done — verified end-to-end (Nyla removed from CloudOps Technologies cleanly) |
| 4 | Termination scope §7.6 (remove all licenses) | Done — verified license count went 1→0 |
| 4 | Termination scope §7.7+§7.8 (mailbox action dispatch — Forward then delete / Delete immediately / Convert to Shared) | Done — all 3 paths verified |
| 4 | Termination scope §7.9 (finalize audit row) | Done — Status branching works for Succeeded/Partial |
| 4 | Azure Automation runbook integration for Convert to Shared | Done — runbook invoked via `shared_azureautomation` connector. Both initial design caveats (license-ordering, runbook-failure-surfacing) fixed and verified. |
| 4 | Critical-error chain consistency for §6.3 GetManager, §6.3 SetManager, §6.4 GetSkus | Done — added Update_AuditRow + Terminate chains so any of these failing properly finalizes the audit row (Status=Failed, EndTimestamp, full StepsFailed/ErrorDetails) instead of cascading to Skipped with no finalization. Also fixed Append_ErrorDetails_GetManager from SetVariable to AppendToStringVariable. |
| 5 | "How it works" walkthrough written | In progress |

## Documents

- [docs/01-sharepoint-schema.md](docs/01-sharepoint-schema.md) — Microsoft Form question list and `Lifecycle Audit Log` list schema
- [docs/02-power-automate-flows.md](docs/02-power-automate-flows.md) — Step-by-step Power Automate flow design for both branches
- [docs/03-graph-api-reference.md](docs/03-graph-api-reference.md) — Exact Graph endpoints, payloads, and required scopes

## Architecture

```
                   +--------------------------+
                   | Admin submits MS Form    |
                   +-------------+------------+
                                 |
                   +-------------v------------+
                   | Power Automate flow      |
                   | Flow-M365-Lifecycle-Main |
                   +-------------+------------+
                                 |
        +------------------------+------------------------+
        |                                                 |
+-------v--------+                              +---------v--------+
| Switch:        |                              | Audit row created|
| Action Type    |---------- references ------->| in SharePoint    |
+----+------+----+                              | (Status=In Prog) |
     |      |                                   +---------+--------+
     |      |                                             |
+----v--+ +-v------+                                      |
| New   | | Term-  |                                      |
| Hire  | | inate  |                                      |
+--+----+ +---+----+                                      |
   |          |                                           |
   |          |   +-------------------+                   |
   +----+-----+---| Microsoft Graph   |                   |
        |         | v1.0 endpoints    |                   |
        |         +-------------------+                   |
        |                                                 |
        |         +-------------------+                   |
        +---------| Office 365        |                   |
        |         | Outlook connector |                   |
        |         +-------------------+                   |
        |                                                 |
        |         +-------------------+                   |
        +---------| Azure Automation  |                   |
                  | runbook (shared   |                   |
                  | mbx conversion)   |                   |
                  +-------------------+                   |
                            |                             |
                            +-----------+-----------------+
                                        |
                              +---------v---------+
                              | Audit row updated |
                              | (Status=Succeeded |
                              | / Partial / Fail) |
                              +-------------------+
```

The flow has one trigger (Microsoft Forms), one switch (action type), two main branches, one finalize scope, and one top-level error handler. Every Graph call has a parallel error-handler block that appends to step-tracking variables flushed to the audit row at the end. The Do until loop around the high-throughput steps honors Graph's `Retry-After` header on HTTP 429.

## How it works

**New Hire path (verified end-to-end through §6.5):**

1. Form submission triggers the flow with the new hire's details: first/last name, job title, department, manager UPN, license SKU choice, groups.
2. The flow initializes nine variables for step tracking and audit, then creates an audit row in `LifecycleAuditLog` with Status=`In Progress`. The row's SharePoint item ID is captured into `varAuditRowID` for later updates.
3. A Switch routes on Action Type. The New Hire case:
   - Composes the UPN, mail nickname, and display name from form fields.
   - **HTTP_CreateUser** — POST `/users` to Graph with a temp password (force-change-on-first-sign-in).
   - **HTTP_GetManager** — GET `/users/{managerUPN}?$select=id` to resolve the manager's object ID.
   - **HTTP_SetManager** — PUT `/users/{newUserId}/manager/$ref` with the manager's `@odata.id`.
   - **HTTP_GetSkus** → Filter array by `skuPartNumber` → Compose first match's `skuId`.
   - **Condition_SkuMissing** — if SKU lookup is empty, terminate the run with `LicenseNotFound`; else proceed to license assignment.
   - **HTTP_AssignLicense** — POST `/users/{newUserId}/assignLicense`.
4. §6.6 (group membership), §6.7 (welcome email), §6.8 (finalize audit row to Status=`Succeeded`) are next.

Every Graph call has a parallel error-handler branch that appends to `varStepsFailed` / `varErrorDetails` on failure. Critical-path failures (CreateUser, license not found, assign license) write a Failed audit row and call `Terminate` to short-circuit the rest of the run.

**Termination path (verified end-to-end, all three mailbox-action choices):**

1. Form submission triggers the same flow; Switch routes Action Type = `Termination` into the termination branch.
2. **§7.1 HTTP_GetTargetUser** — GET `/users/{Q14}?$select=id,displayName,assignedLicenses,userPrincipalName`. Captures user object ID into `varTargetUserObjectID`.
3. **§7.2 Delay Until** — `Compose_TargetTimestamp` evaluates `if(toLower(trim(Q16)) == 'immediate', utcNow(), formatDateTime(Q15 + 'T' + Q16 + ':00'))`. Strict validation: any malformed input (typo like `Immediiate` or gibberish like `xyz`) fails the Compose, terminates with `InvalidTerminationTime`.
4. **§7.3 HTTP_BlockSignIn** — PATCH `/users/{id}` with `{"accountEnabled": false}`. Critical (terminates on failure).
5. **§7.4 HTTP_RevokeSessions** — POST `/users/{id}/revokeSignInSessions` invalidates all refresh tokens. Combined with §7.3 the user is fully ejected.
6. **§7.5 Remove from all groups** — Paginated DoUntil loop fetches `/users/{id}/memberOf` and follows `@odata.nextLink` until empty (handles up to 10,000 memberships defensively). Then Apply_to_each membership: filter to `#microsoft.graph.group`, DELETE `/groups/{id}/members/{userId}/$ref`. Group display names accumulated into `varGroupsRemoved`. Failures are non-critical (logged, flow continues — handles dynamic groups and on-prem-synced groups that reject manual removal).
7. **§7.6 Remove all licenses** — `Select_LicenseSkuIds` maps the cached `assignedLicenses` from §7.1 to an array of `skuId` strings. If empty, skip. Otherwise POST `/users/{id}/assignLicense` with `{addLicenses: [], removeLicenses: <array>}`.
8. **§7.7+§7.8 Mailbox action dispatch** — Switch on Q18:
   - `Forward then delete`: PATCH `/users/{id}/mailboxSettings` with auto-reply pointing to Q19 delegate UPN → POST `/mailFolders/inbox/messageRules` to create a forwarding rule → DELETE `/users/{id}`. Auto-reply and forward rule are non-critical (logged on failure).
   - `Delete immediately`: just DELETE `/users/{id}`.
   - `Convert to Shared`: currently a placeholder — logs `Manual: convert mailbox to shared (Azure Automation runbook not yet wired)` to `varStepsFailed`. Account preserved for manual completion via Exchange Online. To productionize: add `shared_azureautomation` connection and call the `Convert-MailboxToShared` runbook.
9. **§7.9 Finalize audit row** — Same pattern as §6.8 but with termination fields: `Status = if(empty(varStepsFailed), 'Succeeded', 'Partial')`, `MailboxConverted = equals(Q18, 'Convert to Shared')`, `TerminationReason` from Q17, `GroupsRemoved` from varGroupsRemoved.

**Design choice — one audit row per request, not per step.** The audit row is created up front with `In Progress`, mutated through the run, and finalized at the end. Step-level detail lives in the `StepsCompleted` and `ErrorDetails` text columns. This keeps the audit list query-friendly (one row per Form submission) instead of fragmenting into a row per HTTP call. Step granularity is preserved through the flushed string variables.

## Tech stack

- Microsoft 365 E5 (developer tenant)
- Microsoft Forms (trigger)
- Power Automate cloud flow (orchestration) with HTTP, SharePoint, Office 365 Outlook, and Azure Automation connectors
- SharePoint Online (audit log)
- Microsoft Graph v1.0 (identity, group, license, mailbox operations)
- Entra ID app registration with client-credentials OAuth (application permissions only)
- Azure Automation PowerShell 7.2 runbook with the ExchangeOnlineManagement module (shared mailbox conversion via system-assigned managed identity)

No paid third-party tools. No custom Power Automate connectors. Premium Power Automate connectors (HTTP, Azure Automation) included with the Developer Program E5 license.

## Prerequisites

Before building, the tenant needs the following to be in place. Some of these require Global Admin to assign.

### Tenant

- Microsoft 365 E5 (or equivalent SKU that includes Entra ID P1+, Exchange Online, and SharePoint)
- An Azure subscription linked to the same tenant (the free trial is sufficient, used only for Azure Automation)

### Identity

- Entra ID app registration `M365-Lifecycle-Automation` with a client secret. Full setup steps and the six required Graph application permissions are in [02-power-automate-flows.md section 2](docs/02-power-automate-flows.md#2-prerequisites).
- Admin consent granted for all six permissions.

### Azure Automation

- Resource group `rg-m365-automation` in the region closest to your tenant.
- Automation account `aa-m365-lifecycle` with a system-assigned managed identity.
- `ExchangeOnlineManagement` PowerShell module **version 3.4.0** (not the latest) imported into the automation account at runtime **7.2**. Versions 3.6.0 and newer fail with `HRESULT 0x80131047` in the Azure Automation PS 7.2 sandbox due to a .NET 8 dependency the sandbox cannot satisfy. The portal does not let you pick a version, so download the zip locally and upload it. Full command sequence is in [02-power-automate-flows.md section 2.4](docs/02-power-automate-flows.md#24-azure-automation-account-only-needed-if-any-termination-will-choose-convert-to-shared-mailbox--yes) under "Module version pinning".
- Runbook `Convert-MailboxToShared` (PowerShell 7.2) published, code is in [02-power-automate-flows.md section 2.4](docs/02-power-automate-flows.md#24-azure-automation-account-only-needed-if-any-termination-will-choose-convert-to-shared-mailbox--yes).

> **IMPORTANT — managed identity needs TWO grants, not one.** Both must be in place before the `Convert-MailboxToShared` runbook can call `Set-Mailbox`:
>
> 1. **Exchange Recipient Administrator** directory role in Entra. Entra admin center → Identity → Roles & admins → All roles → Exchange Recipient Administrator → Add assignments → search by the managed identity's object ID.
>
> 2. **Exchange.ManageAsApp** application permission on the Office 365 Exchange Online API service principal. The portal does not surface this for managed identities — assign it via the PowerShell snippet in [02-power-automate-flows.md section 2.4](docs/02-power-automate-flows.md#24-azure-automation-account-only-needed-if-any-termination-will-choose-convert-to-shared-mailbox--yes) under "Granting Exchange.ManageAsApp".
>
> Both require Global Admin to grant. Without the API permission, `Connect-ExchangeOnline -ManagedIdentity` returns `UnAuthorized (UnAuthorized)`. The directory role alone is not enough.

### Power Automate

- Connections created for: Microsoft Forms, SharePoint, Office 365 Outlook, HTTP (no pre-built connection, configured per action), Azure Automation.
- Premium connector entitlement (HTTP and Azure Automation are premium, included in the Developer Program E5 tenant).

### SharePoint

- A site to host the audit log. Either the default site or a dedicated site collection at `/sites/ITAutomation`. Record the URL once chosen.
- `Lifecycle Audit Log` list created per [01-sharepoint-schema.md](docs/01-sharepoint-schema.md), including the 28 columns, 5 indexes, and 7 views.
- A second Entra ID app registration used by `scripts/Setup-SharePointList.ps1` for PnP PowerShell authentication. This is distinct from the `M365-Lifecycle-Automation` app the Power Automate flow uses for Graph calls. PnP.PowerShell 2.x removed the bundled multi-tenant Management Shell app and the `PNPPOWERSHELL_CLIENTID` environment variable is not honored in current builds, so the setup script requires the client ID as an explicit parameter.

    Register the app (one line, requires Global Admin and an interactive consent prompt):

    ```powershell
    Register-PnPEntraIDApp -ApplicationName "PnP-Lifecycle-Setup" -Tenant <yourtenant>.onmicrosoft.com -Interactive
    ```

    Record the returned ClientId in the "Tenant-specific values to capture" table below and pass it to the setup script via `-ClientId`.

### Microsoft Forms

- Form `M365 User Lifecycle Request` created per [01-sharepoint-schema.md](docs/01-sharepoint-schema.md) with all 20 questions, branching rules, and the "Record name" setting on.

## Tenant-specific values to capture

These are unique to your tenant. Capture them in a local `config.local.json` (gitignored) — **do not commit them to the public repo**. A `config.example.json` with placeholders is checked in.

| Value | Where to find it |
|---|---|
| Tenant primary domain | Entra → Overview → Primary domain |
| Tenant ID | Entra → Overview |
| App registration client ID (`M365-Lifecycle-Automation`, used by the flow) | App registration overview |
| App registration client secret (`M365-Lifecycle-Automation`) | Stored in Power Automate connection only, never in repo |
| App registration client ID (`PnP-Lifecycle-Setup`, used by the setup script) | Output of `Register-PnPEntraIDApp`, also visible in Entra → App registrations |
| SharePoint audit site URL | After site creation |
| `LifecycleAuditLog` list URL and list GUID | After list creation |
| Form ID | Forms URL after form creation. Pick from dropdown in Power Automate. |
| Power Platform Environment ID + Environment URL | https://admin.powerplatform.microsoft.com → Environments |
| Flow GUID | Power Automate → flow URL `…/flows/{guid}/details` |
| License SKU IDs available | `Get-MgSubscribedSku` output |
| Automation account subscription, resource group, name | Azure portal |

## Test log

The flow's six test scenarios are defined in [02-power-automate-flows.md section 10](docs/02-power-automate-flows.md#10-testing-plan). This table records each test's run outcome and the audit row produced.

| # | Scenario | Date run | Status | Notes |
|---|---|---|---|---|
| 1a | New hire, full success (Jurian Timber) | 2026-05-20 | Succeeded | First full §6.1–§6.5 pass after fixing Set_varAuditRowID, GetManager auth, SetManager body, SetManager runAfter, char(10) → decodeUriComponent('%0A'), and granting Graph Application permissions. User created, manager wired, license assigned. |
| 1b | New hire, partial (Jenner) | 2026-05-20 | Failed (license unassigned) | Pre-fix run. User created in Entra but SetManager success-path action was wired to Failed runAfter, so HTTP_GetSkus and downstream license-assign were skipped. Fixed in subsequent build. |
| 1c | New hire, intermediate runs (Naomi Tucker, James Bond, Jennifer Aniston) | 2026-05-20 | Mixed | Various while iterating on Set_varAuditRowID and Graph permissions. |
| 1d | New hire, §6.6 groups loop (Fatoumata Diawara) | 2026-05-21 | Succeeded | Both group-found path (`CloudOps Technologies` → HTTP_AddGroupMember succeeded) and group-not-found path (`TESTGROUP_DOESNOTEXIST` → logged to varStepsFailed) exercised in same run. Also validated trim() fix for whitespace-padded First Name input. |
| 1e | New hire, §6.7 welcome email (Angelina Carlos) | 2026-05-21 | Succeeded | Send_email_WelcomeNewHire delivered HTML body to Q11 Personal email; rendered correctly with UPN, temp password, sign-in URL, role details, groups. |
| 1f | New hire, §6.8 finalize success (Nyla Brown) | 2026-05-21 | Succeeded | All 7 expected step lines captured in StepsCompleted (CreateUser → GetManager → SetManager → GetSkus → AssignLicense → WelcomeEmail → NewHire branch complete). Status=Succeeded, StepsFailed empty. Confirmed SetVariable→AppendToStringVariable fix. |
| 1g | New hire, §6.8 finalize partial (Mark Anthony) | 2026-05-21 | Succeeded (Status=Partial) | Submitted with one real group + one fake group. Audit row Status=Partial via `if(empty(varStepsFailed),'Succeeded','Partial')`. StepsFailed captured "Group not found". GroupsAdded captured "CloudOps Technologies". |
| 4a | Termination §7.1 resolve target user (Naomi Tucker) | 2026-05-21 | Succeeded | HTTP_GetTargetUser returned id, displayName, UPN, assignedLicenses. Audit row updated with target user details. |
| 4b | Termination §7.2 invalid format (typo `xyz`) | 2026-05-21 | Failed | Compose_TargetTimestamp rejected `2026-05-20Txyz:00`, triggered InvalidTermTime error chain, audit row Status=Failed. Strict-validation path verified. |
| 4c | Termination §7.3 block sign-in (Mark Anthony) | 2026-05-21 | Succeeded | PATCH /users/{id} `{accountEnabled:false}` → 204. Verified Mark's accountEnabled = False in Entra. |
| 4d | Termination §7.4 revoke sessions (Jurian Timber) | 2026-05-21 | Succeeded | POST /revokeSignInSessions → 200 `{value:true}`. Verified signInSessionsValidFromDateTime updated to ~now. |
| 4e | Termination §7.5 remove from groups (Nyla Brown) | 2026-05-21 | Succeeded | DoUntil paginated memberOf, Apply_to_each removed from `CloudOps Technologies`. Other 14 members preserved. |
| 4f | Termination §7.6+§7.8 Convert to Shared (Nyla Brown, resubmit) | 2026-05-21 | Status=Partial | Full §7.1-§7.9 chain. License removed (SPB), groups removed, Switch routed to Convert to Shared placeholder which logs "manual completion required". User account preserved. |
| 4g | Termination §7.8 Delete immediately (James Bond) | 2026-05-21 | Succeeded | Full chain through DELETE /users/{id}. James Bond removed from Entra (verified 404). |
| 4h | Termination §7.8 Forward then delete (Jennifer Aniston) | 2026-05-21 | Partial | SetAutoReply + CreateForwardRule both got 404 (mailbox not provisioned for this test user in dev tenant — design correctly marks these R=non-blocking). DeleteUser succeeded, Jennifer removed from Entra. |
| 4i | Termination §7.8 Convert to Shared via Azure Automation runbook (Nyla Brown resubmit) | 2026-05-21 | Succeeded (PA) / Failed (runbook internally) | Runbook invocation verified end-to-end: MI auth, ExchangeOnlineManagement v3.4.0 module, Connect-ExchangeOnline, Set-Mailbox all working. Runbook itself failed because Nyla's mailbox was already gone (her license was removed in run 4f, deleting the mailbox). PA initially reported Succeeded because the connector signals on job *creation*, not job *outcome*. Two design caveats then identified — both resolved in 4j below. |
| 4j | Convert to Shared caveat fixes verified (Nyla Brown resubmit) | 2026-05-21 | Succeeded (run) / Status=Partial (audit) | Both caveats from 4i now fixed in the flow. License removal correctly skipped for Convert-to-Shared (Condition_HasLicenses requires Q18 != 'Convert to Shared'). Runbook failure correctly surfaced via Condition_RunbookCompleted gate — Append_StepsCompleted_ConvertToShared was Skipped, Append_StepsFailed_RunbookExecution + Append_ErrorDetails_RunbookExecution fired with full job exception text. Audit row final Status=Partial reflects the real runbook failure instead of a misleading Succeeded. |
| 5 | Force 429 with rapid submissions | — | Not run | Worth running once. Validates retryPolicy on HTTP actions. |

## Tooling: edit the flow as JSON, not in the designer

Mid-build I pivoted from clicking through the Power Automate v3 designer to editing the flow's JSON definition directly via the Power Automate Management REST API. The v3 designer's Code view paste reverted silently, the action picker stuck after a delete, and several saves quietly dropped changes. The REST API gives a clean round-trip.

Scripts in the repo root:

| Script | Purpose |
|---|---|
| `Get-FlowDefinition.ps1` | GET the flow JSON to a local file. Resolves display name → GUID via list call. |
| `Set-FlowDefinition.ps1` | PATCH only `properties.definition` (and `connectionReferences`) back to the env. Preserves runtime state and connections. |
| `Grant-AppGraphPermissions.ps1` | Add Microsoft Graph Application permissions to an Entra app reg and grant admin consent in one shot. |
| `scripts/Setup-SharePointList.ps1` | One-shot PnP provisioning of the audit list per `docs/01-sharepoint-schema.md`. |
| `scripts/New-TestUsers.ps1` | Seeds the dev tenant with five throwaway test users. |

Typical edit cycle:

```powershell
# 1. Pull
.\Get-FlowDefinition.ps1 -EnvironmentId <env> -FlowName "Flow-M365-Lifecycle-Main"

# 2. Edit flow-Flow-M365-Lifecycle-Main.json in VS Code

# 3. Push
.\Set-FlowDefinition.ps1 -EnvironmentId <env> -FlowName <flow-guid> -DefinitionPath .\flow-Flow-M365-Lifecycle-Main.json -Force
```

The local JSON contains the flow's embedded Graph client secret. The pattern `flow-*.json` is gitignored — never commit it.

## Lessons learned

- **Power Automate v3 designer is too fragile for non-trivial edits.** Pasting JSON into Code view often reverts silently. The fix was to drive the flow's REST API directly via PowerShell. That makes edits source-controllable and reproducible.
- **`char(10)` is an Excel function, not a Power Automate Workflow Definition Language function.** Use `decodeUriComponent('%0A')` for newline. Every Append-to-string action also requires the `@` prefix on the expression — without it the literal text `concat(...)` gets appended instead of evaluating.
- **App registrations created with only Delegated Graph permissions cannot do client-credentials auth.** Flow HTTP actions run as the app itself with no user context. The app needs **Application** type permissions (`User.ReadWrite.All`, `Organization.Read.All`, `Group.ReadWrite.All`) granted with admin consent. Six pre-existing Delegated permissions on the same app reg are unused.
- **M365 Developer Program tenants don't pre-provision Dataverse**, so `pac` CLI / solution-based workflows don't work out of the box. Bootstrapping a tenant Global Admin into a Dataverse System Administrator role is also blocked by a chicken-and-egg around `prvAssignRole` and a known bug in `pac admin self-elevate` (sends an empty `api-version=` parameter in pac 2.7.4). For this project, none of that matters — the Power Automate REST API path doesn't need Dataverse at all.
- **`Update item` against a SharePoint list needs the row's numeric ID, not a string.** The audit row's ID has to be captured with `body('Create_AuditRow')?['ID']` into a properly typed Integer variable. Wrapping it in `string()` produces a runtime type-mismatch error.
- **Production hardening this dev-tenant build deliberately skips:** the flow's HTTP actions embed the client secret instead of using a connection reference; there's no managed identity for the flow itself; no PIM gating on the app reg's role grants; no Sentinel/KQL alerts on the audit list; no idempotency key on the Form trigger. All called out in [docs/02-power-automate-flows.md §11](docs/02-power-automate-flows.md).

## Known limitations (resolved)

Two design issues were uncovered while wiring §7.8 Convert to Shared via Azure Automation, and both have been fixed:

**A. License removal vs mailbox conversion ordering — RESOLVED.** §7.6 `Condition_HasLicenses` now skips the license-removal block when Q18 = `Convert to Shared`. The mailbox stays intact for the runbook. For the other mailbox actions (Forward then delete, Delete immediately), license removal still runs before user deletion as before.

**B. Runbook failures not surfaced — RESOLVED.** Added `Condition_RunbookCompleted` inside `Case_ConvertToShared` that gates the success-append on `body('Create_Job_ConvertToShared')?['properties']?['status'] == 'Completed'`. The else branch writes a CRITICAL marker to varStepsFailed and captures the full job exception in varErrorDetails. §7.9 Finalize then correctly lands Status=Partial (or Failed) when the runbook didn't complete cleanly. Previously misleading "Succeeded" outcomes on runbook failures are gone — verified 2026-05-21 with a re-run that correctly reported Status=Partial and the full Exchange Online error.

## Out of scope

- HRIS integration. Triggering is manual Form submission only.
- Cross-tenant scenarios. Single tenant.
- Production hardening (no managed-identity flow context, no Privileged Identity Management gates, no KQL alerts).
- Custom connectors. Built-in connectors only.

## Disclaimer

This project is a proof-of-concept built in a Microsoft 365 Developer Program tenant and is provided as-is, without warranty, for educational and reference purposes only. It is not production-hardened. Specifically:

- The Microsoft Graph client secret is embedded inline in the Power Automate flow JSON. The flow JSON is gitignored locally and never committed, but the production-correct pattern is to externalize secrets to Azure Key Vault and reference them via a managed identity.
- No automated secret-rotation is wired up. The 6-month expiry is tracked manually.
- The SharePoint, Office 365 Outlook, Microsoft Forms, and Azure Automation connections authenticate as a user identity. If that user leaves the tenant, the flow breaks.
- All tenant identifiers (site URLs, list GUIDs, Form question IDs, app object IDs, license SKU IDs, automation account names) are tenant-specific. They will not work in another tenant without re-provisioning every artifact and updating the references in the flow JSON.
- The Azure Automation runbook used for Convert-to-Shared assumes Exchange Online PowerShell module 3.x and managed-identity-on-the-automation-account auth; both must be verified in any new tenant.

Before adapting any pattern from this repo for a production tenant, perform a security review and replace each of the above with the production-correct alternative.

## License

MIT. See [LICENSE](LICENSE).
