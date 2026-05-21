# Deliverable 1: Microsoft Form and SharePoint Audit Log Schema

This document defines the two data structures that the lifecycle automation depends on:

1. The Microsoft Form that triggers the Power Automate flow
2. The SharePoint list "Lifecycle Audit Log" that records every flow run

Build them in this order. The flow in Deliverable 2 binds to the column internal names defined here, so renaming columns after the flow is built will break dynamic-content references.

---

## 1. Microsoft Form: "M365 User Lifecycle Request"

### Form-level settings

| Setting | Value | Rationale |
|---|---|---|
| Who can fill out this form | Specific people in my organization | Restrict to IT admins until production-ready |
| Record name | On | Auto-captures submitter UPN and display name, no need to ask for them as questions |
| One response per person | Off | Same admin submits multiple requests |
| Accept responses | On | Required for the flow trigger to fire |
| Response receipts | Off | The flow sends a custom confirmation email |
| Customize thank you message | "Request received. You will get a confirmation email when the automation completes." | Sets expectations |

When "Record name" is on, the Forms connector in Power Automate exposes two extra dynamic-content fields on the trigger: `Responder's Email` and `Name`. The flow uses both for audit logging without you adding submitter questions.

### Question list

The form uses Microsoft Forms branching. Q1 is the gate. After Q1 is answered, the responder sees either the New Hire branch (Q2 to Q12) or the Termination branch (Q13 to Q20), never both.

| # | Question text | Field type | Required | Options / format | Notes |
|---|---|---|---|---|---|
| 1 | Action type | Choice (single answer) | Yes | New Hire, Termination | Branching source. Configure: if "New Hire" go to Q2, if "Termination" go to Q13. |
| 2 | First name | Text (short answer) | Yes | Letters, hyphens, apostrophes | Used to build mailNickname and displayName |
| 3 | Last name | Text (short answer) | Yes | Letters, hyphens, apostrophes | Used to build mailNickname and displayName |
| 4 | Preferred display name (if different from "First Last") | Text (short answer) | No | Free text | If blank, flow constructs "First Last" |
| 5 | Job title | Text (short answer) | Yes | Free text | Maps to Graph `jobTitle` |
| 6 | Department | Choice (single answer) | Yes | IT, Finance, HR, Operations, Engineering, Sales, Executive | Populate with your real department list. Maps to Graph `department`. |
| 7 | Office location | Choice (single answer) | Yes | HQ, Remote, Field, Other | Maps to Graph `officeLocation` |
| 8 | Manager email (UPN) | Text (short answer) | Yes | Restrict to "Email" type | Maps to Graph manager reference via `/users/{id}/manager/$ref` |
| 9 | Start date | Date | Yes | Date picker | Used in welcome email, not in any Graph call (Entra ID has no native start-date attribute on the user object) |
| 10 | License SKU to assign | Choice (single answer) | Yes | List the SkuPartNumber values from your tenant (see "Looking up your license SKU IDs" below) | The flow translates the SkuPartNumber to a SkuId GUID via a lookup step |
| 11 | Distribution and security groups to add (one per line, by display name) | Text (long answer) | No | One group per line | Flow splits on newline, resolves each name to an objectId, adds via `/groups/{id}/members/$ref` |
| 12 | Mobile phone number | Text (short answer) | No | E.164 format e.g. +12025551234 | Maps to Graph `mobilePhone` |
| 13 | User to offboard (UPN) | Text (short answer) | Yes | Restrict to "Email" type | The flow resolves this to an objectId via `/users/{upn}` |
| 14 | Effective termination date and time (tenant local time) | Date | Yes | Date picker. Forms does not natively offer a combined date-time picker, so add Q15 below for time. | Combined with Q15 in the flow |
| 15 | Termination time of day | Choice (single answer) | Yes | 09:00, 12:00, 17:00, End of business day (17:00), Immediate | "Immediate" tells the flow to run all termination steps now instead of waiting |
| 16 | Convert mailbox to shared after termination | Choice (single answer) | Yes | Yes, No | If Yes, the flow calls `Set-Mailbox -Type Shared` via the Exchange Online connector |
| 17 | Forward incoming mail to a delegate | Choice (single answer) | Yes | Yes, No | Drives branching to Q18 |
| 18 | Delegate UPN to receive forwarded mail | Text (short answer) | Conditional (required when Q17 is Yes) | Restrict to "Email" type | Set Forms branching: Q17 = Yes goes to Q18, Q17 = No skips to Q19 |
| 19 | Mailbox retention period | Choice (single answer) | Yes | 30 days, 60 days, 90 days, Indefinite | Logged for audit. Actual retention is governed by the Microsoft 365 retention policy, the flow only records the requested value. |
| 20 | Reason for termination | Choice (single answer) | Yes | Voluntary resignation, Involuntary, Retirement, Contract end, Other | Recorded for audit only |

### Branching configuration (exact clicks in Forms)

1. Open the form, click the "..." menu on Q1, choose "Add branching".
2. On the branching screen, for Q1 "Action type":
   - Set "New Hire" branch target to Q2.
   - Set "Termination" branch target to Q13.
3. On Q12 (last New Hire question), set "Go to" to "End of the form".
4. On Q17, set:
   - "Yes" branch target to Q18.
   - "No" branch target to Q19.
5. Q20 should have "Go to" set to "End of the form".

### Looking up your license SKU IDs

Power Automate needs the SkuId GUID, not the friendly name, to call `/users/{id}/assignLicense`. To get the list of available SKUs in your tenant, run one of these:

PowerShell with Microsoft Graph SDK:

```powershell
Connect-MgGraph -Scopes "Organization.Read.All"
Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId, ConsumedUnits, @{n="Enabled";e={$_.PrepaidUnits.Enabled}} | Format-Table -AutoSize
```

Or a direct Graph call (in Graph Explorer at https://developer.microsoft.com/en-us/graph/graph-explorer):

```
GET https://graph.microsoft.com/v1.0/subscribedSkus?$select=skuPartNumber,skuId,consumedUnits,prepaidUnits
```

Take the `SkuPartNumber` column from the output and use those strings as the Q10 choice options. Common values in an E5 developer tenant include `ENTERPRISEPREMIUM` (E5), `ENTERPRISEPACK` (E3), `EMSPREMIUM` (EMS E5), and `FLOW_FREE`. Use what your tenant actually returns, not this list.

---

## 2. SharePoint list: "Lifecycle Audit Log"

### List-level settings

| Setting | Value |
|---|---|
| Site | Use the default site of your dev tenant, or create a dedicated site collection at `/sites/ITAutomation`. Document the URL in your README. |
| List name | `LifecycleAuditLog` (no spaces in the internal URL segment, easier to reference in Power Automate) |
| Display name | `Lifecycle Audit Log` (set after creation, this only changes the display, not the URL) |
| Versioning | On, keep last 50 major versions (one row gets multiple status updates during a single flow run) |
| Content approval | Off |
| Quick edit | On (useful for manual audit corrections) |
| Item-level permissions | Default. Tighten before production, leave open in dev. |

### Important note about column internal names

When you create a SharePoint column with spaces in the display name (for example "Action Type"), SharePoint encodes the internal name as `Action_x0020_Type`. The internal name is what Power Automate dynamic content uses, and you cannot rename it once created.

Create every column below with the internal name in the "Column name" field first (no spaces, PascalCase), save the column, then edit it and set the display name with spaces if you want a friendlier label.

### Column schema

All columns below are custom unless noted. The default `Title` column is renamed and reused as the audit record identifier.

| Internal name | Display name | Type | Required | Default | Notes |
|---|---|---|---|---|---|
| Title | Audit Record ID | Single line of text | Yes | (empty) | Renamed from default Title. Format: `NH-yyyyMMdd-HHmmss-{first 8 chars of flow run GUID}` or `TERM-yyyyMMdd-HHmmss-{first 8 chars}`. The flow generates this in its first compose step. |
| ActionType | Action Type | Choice | Yes | (none) | Choices: `New Hire`, `Termination`. Display choices using radio buttons. |
| TargetUserUPN | Target User UPN | Single line of text | Yes | (empty) | Lowercase, validated email format. For new hires, the UPN the flow constructed. For terminations, the UPN from Q13. |
| TargetUserDisplayName | Target User Display Name | Single line of text | No | (empty) | Resolved from Graph after user creation or lookup |
| TargetUserObjectID | Target User Object ID | Single line of text | No | (empty) | Entra ID object GUID. Empty until the user is created (new hire) or looked up (termination). |
| SubmitterEmail | Submitter Email | Single line of text | Yes | (empty) | From Forms "Responder's Email" dynamic content |
| SubmitterDisplayName | Submitter Display Name | Single line of text | No | (empty) | From Forms "Name" dynamic content |
| FormResponseID | Form Response ID | Single line of text | Yes | (empty) | Forms response GUID, useful for cross-referencing if a row is questioned |
| FlowRunID | Flow Run ID | Single line of text | Yes | (empty) | `workflow()['run']['name']` from the flow expression library |
| FlowRunURL | Flow Run URL | Hyperlink | No | (empty) | Constructed in the flow, lets you click from audit row straight to the run history |
| StartTimestamp | Start Timestamp | Date and Time | Yes | (none) | "Date and Time" format, "Standard" friendly display. UTC. |
| EndTimestamp | End Timestamp | Date and Time | No | (none) | Set by the final SharePoint Update Item step |
| DurationSeconds | Duration (seconds) | Number | No | 0 | Integer. Calculated as `ticks(EndTimestamp) - ticks(StartTimestamp)` divided by 10,000,000. |
| Status | Status | Choice | Yes | In Progress | Choices: `In Progress`, `Succeeded`, `Partial Success`, `Failed`. The row is created with In Progress, updated as the flow runs. |
| StepsCompleted | Steps Completed | Multiple lines of text (plain text) | No | (empty) | Newline-separated list of step names that succeeded. The flow appends to this string at each step. |
| StepsFailed | Steps Failed | Multiple lines of text (plain text) | No | (empty) | Newline-separated list of step names that failed |
| ErrorDetails | Error Details | Multiple lines of text (plain text) | No | (empty) | JSON-formatted dump of any error response bodies. Plain text, not rich text, so you can parse it. |
| RetryCount | Retry Count | Number | No | 0 | Total number of HTTP 429 retries across all Graph calls in this run |
| LicenseSKU | License SKU | Single line of text | No | (empty) | SkuPartNumber from Q10. Empty for terminations. |
| LicenseSKUID | License SKU ID | Single line of text | No | (empty) | SkuId GUID actually assigned |
| Department | Department | Choice | No | (none) | Same choice list as Forms Q6. Add `(none)` as a choice so it can be cleared for terminations. |
| ManagerUPN | Manager UPN | Single line of text | No | (empty) | From Q8 (new hire). Empty for terminations unless mailbox forwarding was set, in which case the delegate UPN goes in MailboxForwardingTo. |
| StartDate | Start Date | Date Only | No | (none) | New hires only |
| TerminationDate | Termination Date | Date and Time | No | (none) | Terminations only. Composed from Q14 and Q15. |
| MailboxConverted | Mailbox Converted to Shared | Yes/No | No | No | True if the flow ran the conversion. False for new hires and for terminations where Q16 was No. |
| MailboxForwardingTo | Mailbox Forwarding To | Single line of text | No | (empty) | Delegate UPN from Q18 if Q17 was Yes |
| GroupsAdded | Groups Added | Multiple lines of text (plain text) | No | (empty) | Newline-separated list of group display names successfully added |
| GroupsRemoved | Groups Removed | Multiple lines of text (plain text) | No | (empty) | For terminations, the groups the user was removed from |
| TerminationReason | Termination Reason | Choice | No | (none) | Same options as Q20. Add `(none)` for new-hire rows. |
| RetentionPeriod | Retention Period | Choice | No | (none) | Same options as Q19 |
| Notes | Notes | Multiple lines of text (rich text) | No | (empty) | Free-form notes added manually after the fact |

### Required flags summary

Set these to required at the column level so SharePoint blocks any flow step that tries to create or update an item without them:

- Title
- ActionType
- TargetUserUPN
- SubmitterEmail
- FormResponseID
- FlowRunID
- StartTimestamp
- Status

Every other column is optional because new-hire rows and termination rows fill different subsets.

### Indexes

SharePoint allows up to 20 indexes per list. The list view threshold of 5,000 items will not be hit in a dev tenant, but indexing the columns used for filtering and sorting now means views will keep working if this is ever ported to a production tenant with high volume.

Create single-column indexes on:

| Column | Reason |
|---|---|
| FlowRunID | Unique lookup when correlating a Power Automate run to an audit row |
| TargetUserUPN | Filter "all activity for user X" |
| ActionType | Filter views by New Hire vs Termination |
| Status | Filter Failed and In Progress views |
| StartTimestamp | Default sort key, used in every view |

To create an index: List settings → Indexed columns → Create a new index → select column → OK.

### Recommended views

Create these views via List settings → Views → Create view → Standard view.

| View name | Filter | Sort | Columns shown | Purpose |
|---|---|---|---|---|
| All Items (default) | (none) | StartTimestamp descending | Title, ActionType, TargetUserUPN, Status, StartTimestamp, SubmitterEmail, FlowRunURL | The default landing view |
| In Progress | Status equals `In Progress` | StartTimestamp descending | Title, ActionType, TargetUserUPN, StartTimestamp, FlowRunURL | Spot stuck or long-running flows |
| Failed | Status equals `Failed` OR Status equals `Partial Success` | StartTimestamp descending | Title, ActionType, TargetUserUPN, Status, StepsFailed, ErrorDetails, FlowRunURL, RetryCount | Triage queue |
| New Hires (Last 30 Days) | ActionType equals `New Hire` AND StartTimestamp is greater than `[Today]-30` | StartTimestamp descending | Title, TargetUserUPN, TargetUserDisplayName, Department, ManagerUPN, LicenseSKU, Status, StartDate | Onboarding activity report |
| Terminations (Last 30 Days) | ActionType equals `Termination` AND StartTimestamp is greater than `[Today]-30` | StartTimestamp descending | Title, TargetUserUPN, TerminationDate, MailboxConverted, MailboxForwardingTo, TerminationReason, Status | Offboarding activity report |
| By Submitter | (none) | SubmitterEmail ascending, then StartTimestamp descending | Title, ActionType, TargetUserUPN, Status, StartTimestamp | Grouped reporting by who ran what |
| Audit Detail | (none) | StartTimestamp descending | All columns except FlowRunURL and Notes | Full-fidelity export, hide in default view to keep the All Items view scannable |

For the New Hires and Terminations views, the filter syntax in SharePoint is literally `[Today]-30`, typed into the value field. Do not wrap it in quotes.

### Where the flow will create vs update the audit row

The flow creates the audit row at the start of the run (Status = In Progress, StartTimestamp set, most other fields populated from the Form response and constructed values). Each step in the lifecycle appends to StepsCompleted or StepsFailed via an Update Item action. The final step sets Status, EndTimestamp, DurationSeconds, and any closing fields. This pattern is detailed in Deliverable 2.

---

## Build checklist

Run through these before declaring Deliverable 1 done:

- [ ] Microsoft Form created with all 20 questions in the order above
- [ ] Branching configured on Q1, Q17, and Q12 (end-of-form)
- [ ] License SKU options in Q10 populated from your actual tenant `subscribedSkus` output
- [ ] Department options in Q6 reflect your real org structure
- [ ] "Record name" toggled on
- [ ] SharePoint site identified and URL recorded in README
- [ ] `LifecycleAuditLog` list created
- [ ] All 28 columns added with the correct internal names and types
- [ ] Required flags set on the 8 columns listed above
- [ ] 5 indexes created
- [ ] 7 views created
- [ ] Default Title column renamed to "Audit Record ID"
- [ ] Test row created manually and deleted to confirm all column types accept their expected input
