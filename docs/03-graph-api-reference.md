# Deliverable 3: Microsoft Graph API Reference

Every Microsoft Graph call made by `Flow-M365-Lifecycle-Main`, with the exact endpoint, HTTP method, required application permission, request body, expected response, and tenant-specific quirks.

All endpoints use the `v1.0` Graph endpoint. Beta endpoints are noted where used (none currently). Base URL is `https://graph.microsoft.com/v1.0`.

The flow authenticates with OAuth 2.0 client credentials against the app registration created in Deliverable 2 section 2.1. Every request includes:

```
Authorization: Bearer {access_token}
Content-Type: application/json
```

The HTTP action in Power Automate handles token acquisition and refresh internally when configured as described in Deliverable 2 section 4.1.

---

## 1. Permission summary

The app registration needs these six application permissions, all granted admin consent. Every endpoint below lists which one it requires.

| Permission | Endpoints that need it |
|---|---|
| `User.ReadWrite.All` | Create user, get user, update user, block sign-in, revoke sessions, look up manager |
| `Directory.ReadWrite.All` | Assign license, remove license, list subscribed SKUs |
| `Group.ReadWrite.All` | Get group by displayName, add member, remove member, list memberships |
| `Organization.Read.All` | List subscribed SKUs (covered by Directory.ReadWrite.All, listed for least-privilege hardening) |
| `MailboxSettings.ReadWrite` | Set auto-reply via mailboxSettings |
| `Mail.ReadWrite` | Create inbox forwarding rule |

If you ever migrate this to delegated permissions (for a different trigger pattern), the equivalents are `User.ReadWrite`, `Group.ReadWrite.All`, `MailboxSettings.ReadWrite`, `Mail.ReadWrite`. `Directory.ReadWrite.All` has no delegated equivalent that grants the same write scope, you would need separate `User.ManageIdentities.All` plus tenant-admin consent.

---

## 2. Users

### 2.1 Create user

| Property | Value |
|---|---|
| Endpoint | `POST /users` |
| Permission | `User.ReadWrite.All` |
| Flow step | 6.2 `HTTP_CreateUser` |
| Success status | 201 Created |

Request:

```json
{
  "accountEnabled": true,
  "displayName": "Jordan Lee",
  "mailNickname": "jordan.lee",
  "userPrincipalName": "jordan.lee@yourtenant.onmicrosoft.com",
  "passwordProfile": {
    "forceChangePasswordNextSignIn": true,
    "password": "a1b2c3d4e5fA7!"
  },
  "jobTitle": "Systems Administrator",
  "department": "IT",
  "officeLocation": "HQ",
  "mobilePhone": "+12025551234",
  "usageLocation": "US"
}
```

Response (abbreviated):

```json
{
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users/$entity",
  "id": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "userPrincipalName": "jordan.lee@yourtenant.onmicrosoft.com",
  "displayName": "Jordan Lee",
  "mail": null
}
```

The `id` field is the Entra Object ID. The flow stores it in `varTargetUserObjectID`. The `mail` field returns null at creation time, Exchange Online provisions the mailbox asynchronously after license assignment.

Quirks:

- `userPrincipalName` must use a domain that is verified in the tenant. In a dev tenant, that is `{tenantname}.onmicrosoft.com` by default.
- `mailNickname` cannot contain spaces or these characters: `@()\\[]";:<>,SPACE`. The flow strips them in the compose step.
- `usageLocation` is a two-letter ISO 3166-1 alpha-2 code. It must be set at creation or before the first `assignLicense` call, otherwise license assignment returns 400.
- `passwordProfile.password` must comply with the tenant password policy. The expression in Deliverable 2 section 3.3 guarantees the default policy is met.
- The default response includes only a subset of fields. To get more in a single call, append `?$select=id,displayName,userPrincipalName,mail,assignedLicenses` to the URL.

### 2.2 Get user by UPN or ID

| Property | Value |
|---|---|
| Endpoint | `GET /users/{upn-or-id}` |
| Permission | `User.ReadWrite.All` (or `User.Read.All`) |
| Flow step | 6.3 `HTTP_GetManager`, 7.1 `HTTP_GetTargetUser` |
| Success status | 200 OK |

URL examples:

```
GET /users/jordan.lee@yourtenant.onmicrosoft.com
GET /users/a1b2c3d4-e5f6-7890-abcd-ef1234567890
GET /users/jordan.lee@yourtenant.onmicrosoft.com?$select=id,displayName,assignedLicenses,userPrincipalName
```

Use `$select` to limit the payload, the default response includes ~25 fields, most are unused.

Quirks:

- UPN lookups are case-insensitive but return 404 if the user does not exist or is soft-deleted.
- Soft-deleted users live at `/directory/deletedItems/microsoft.graph.user/{id}` for 30 days. The flow does not handle this case, a terminated user re-submitted for offboarding will return 404.
- Guest users are reachable here with their full guest UPN (`firstname_externaldomain.com#EXT#@yourtenant.onmicrosoft.com`).

### 2.3 Update user attributes

| Property | Value |
|---|---|
| Endpoint | `PATCH /users/{id}` |
| Permission | `User.ReadWrite.All` |
| Flow step | 7.3 `HTTP_BlockSignIn` |
| Success status | 204 No Content |

Block sign-in request:

```json
{ "accountEnabled": false }
```

Quirks:

- PATCH returns 204 with no body on success. The HTTP action's success branch should not try to parse the response body.
- Setting `accountEnabled` to false does not invalidate existing tokens. Step 7.4 (revoke sessions) is the second half of the lockout.

### 2.4 Revoke all sign-in sessions

| Property | Value |
|---|---|
| Endpoint | `POST /users/{id}/revokeSignInSessions` |
| Permission | `User.ReadWrite.All` |
| Flow step | 7.4 `HTTP_RevokeSessions` |
| Success status | 200 OK |

Request body: none. The HTTP action body field must be left empty, not `{}`.

Response:

```json
{
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Edm.Boolean",
  "value": true
}
```

Quirks:

- This revokes refresh tokens. Access tokens already issued remain valid until their lifetime expires (typically 1 hour, configurable via Conditional Access). To force-evict mid-session, also configure a Conditional Access policy with "Sign-in frequency = every time" for terminated users, outside the scope of this flow.
- The action also revokes Microsoft 365 service tokens for Exchange, SharePoint, and Teams via their respective backchannels.

---

## 3. Manager relationship

### 3.1 Set manager

| Property | Value |
|---|---|
| Endpoint | `PUT /users/{user-id}/manager/$ref` |
| Permission | `User.ReadWrite.All` |
| Flow step | 6.3 `HTTP_SetManager` |
| Success status | 204 No Content |

Request:

```json
{
  "@odata.id": "https://graph.microsoft.com/v1.0/users/{manager-id}"
}
```

The `@odata.id` value is the full Graph URL of the manager user object, not a bare ID.

Quirks:

- Returns 400 if the referenced manager does not exist or is not a user object.
- Setting a manager creates a one-way reference. The manager's `directReports` collection updates automatically, no second call needed.
- To remove a manager: `DELETE /users/{user-id}/manager/$ref`. Not used in this flow.

### 3.2 Get manager

| Property | Value |
|---|---|
| Endpoint | `GET /users/{id}/manager` |
| Permission | `User.Read.All` (or higher) |
| Flow step | Not used, listed for completeness |
| Success status | 200 OK |

Returns the manager user object. Returns 404 if no manager is set.

---

## 4. Groups

### 4.1 Look up group by display name

| Property | Value |
|---|---|
| Endpoint | `GET /groups?$filter=displayName eq '{name}'&$select=id,displayName` |
| Permission | `Group.Read.All` (or higher) |
| Flow step | 6.6 `HTTP_GetGroup` |
| Success status | 200 OK |

Example:

```
GET /groups?$filter=displayName eq 'Engineering Team'&$select=id,displayName
```

For a group name containing an apostrophe like `O'Brien Team`, the filter value must escape the apostrophe by doubling it:

```
GET /groups?$filter=displayName eq 'O''Brien Team'&$select=id,displayName
```

The Power Automate expression that produces this is in Deliverable 2 section 6.6.

Response:

```json
{
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups(id,displayName)",
  "value": [
    {
      "id": "11111111-2222-3333-4444-555555555555",
      "displayName": "Engineering Team"
    }
  ]
}
```

Quirks:

- `value` is an array because display names are not unique. The flow uses `first(...)` and accepts that risk in a dev tenant. In production, filter on `mailNickname` (which is unique within a tenant) instead, or require GUIDs in the form.
- An empty `value` array means no match. The flow's condition `length(body('HTTP_GetGroup')?['value']) greater than 0` covers this.
- Distribution lists, Microsoft 365 groups, security groups, and mail-enabled security groups all appear here. The flow does not differentiate, but `add member` (4.2) fails on certain types listed below.

### 4.2 Add member to group

| Property | Value |
|---|---|
| Endpoint | `POST /groups/{group-id}/members/$ref` |
| Permission | `Group.ReadWrite.All` (or `GroupMember.ReadWrite.All`) |
| Flow step | 6.6 `HTTP_AddGroupMember` |
| Success status | 204 No Content |

Request:

```json
{
  "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/{user-id}"
}
```

Note the path segment is `/directoryObjects/`, not `/users/`. Both work in practice for user references, but `directoryObjects` is the documented form and supports adding nested groups and service principals the same way.

Quirks:

- Returns 400 `One or more added object references already exist` if the user is already a member. The flow treats this as success, the desired state matches.
- Distribution lists and mail-enabled security groups synced from on-premises Active Directory return 403, they are read-only in Graph.
- Dynamic groups return 403, membership is rule-based.
- The `All Users` and `All Company` automatic groups return 403.
- Adding a user to a group that grants a license via group-based licensing assignment will trigger an automatic license assignment, which counts against your tenant license pool.

### 4.3 Remove member from group

| Property | Value |
|---|---|
| Endpoint | `DELETE /groups/{group-id}/members/{user-id}/$ref` |
| Permission | `Group.ReadWrite.All` |
| Flow step | 7.5 `HTTP_RemoveMember` |
| Success status | 204 No Content |

Request body: none.

Quirks:

- Returns 404 if the user is not a member. The flow logs this and continues.
- Same on-premises sync and dynamic-group restrictions as 4.2.
- Removing a user from a license-granting group does not immediately revoke the license, the asynchronous group-licensing job picks it up within a few minutes.

### 4.4 List a user's group memberships

| Property | Value |
|---|---|
| Endpoint | `GET /users/{id}/memberOf?$select=id,displayName` |
| Permission | `Group.Read.All` and `User.Read.All` |
| Flow step | 7.5 `HTTP_GetMemberships` |
| Success status | 200 OK |

Response:

```json
{
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#directoryObjects",
  "value": [
    {
      "@odata.type": "#microsoft.graph.group",
      "id": "11111111-...",
      "displayName": "Engineering Team"
    },
    {
      "@odata.type": "#microsoft.graph.directoryRole",
      "id": "22222222-...",
      "displayName": "User Administrator"
    }
  ]
}
```

Quirks:

- The response mixes group memberships and directory role assignments. The flow filters on `@odata.type` equal to `#microsoft.graph.group` to ignore roles.
- To get only groups, use `/users/{id}/transitiveMemberOf/microsoft.graph.group` instead, which casts the result type. The flow uses the simpler form and filters in the loop.
- Default page size is 100. The flow assumes no user is in more than 100 groups. For users with more memberships, follow the `@odata.nextLink` pagination link, see section 9.

---

## 5. Licenses

### 5.1 List tenant subscribed SKUs

| Property | Value |
|---|---|
| Endpoint | `GET /subscribedSkus?$select=skuId,skuPartNumber,consumedUnits,prepaidUnits` |
| Permission | `Directory.Read.All` (or `Organization.Read.All`) |
| Flow step | 6.4 `HTTP_GetSkus` |
| Success status | 200 OK |

Response:

```json
{
  "value": [
    {
      "skuId": "c7df2760-2c81-4ef7-b578-5b5392b571df",
      "skuPartNumber": "ENTERPRISEPREMIUM",
      "consumedUnits": 3,
      "prepaidUnits": { "enabled": 25, "suspended": 0, "warning": 0 }
    }
  ]
}
```

Quirks:

- The list contains every SKU your tenant has ever had, including trials that expired. Free SKUs like `FLOW_FREE`, `STREAM`, and `POWER_BI_STANDARD` show up here too.
- `consumedUnits` is the current assignment count, `prepaidUnits.enabled` is the cap. The flow does not check headroom before assigning, you should add a condition if running near license cap.
- This endpoint can return a stale response for up to 5 minutes after a SKU change. Not a problem for this flow's read-once pattern.

### 5.2 Assign license

| Property | Value |
|---|---|
| Endpoint | `POST /users/{id}/assignLicense` |
| Permission | `User.ReadWrite.All` plus `Directory.ReadWrite.All` |
| Flow step | 6.5 `HTTP_AssignLicense` |
| Success status | 200 OK |

Request:

```json
{
  "addLicenses": [
    {
      "skuId": "c7df2760-2c81-4ef7-b578-5b5392b571df",
      "disabledPlans": []
    }
  ],
  "removeLicenses": []
}
```

Response: returns the full user object after assignment.

Quirks:

- The user must have `usageLocation` set, otherwise 400 with error code `UserNotLicensable`.
- `disabledPlans` accepts an array of `servicePlanId` GUIDs to disable specific services within the SKU (for example, disable Yammer within E5). Look up service plan IDs at `/subscribedSkus?$select=skuPartNumber,servicePlans`. Empty array means all services enabled.
- Cannot use this endpoint to assign a license that comes from group-based licensing, those assignments are managed at the group level.
- Assignments are not atomic. If `addLicenses` contains two SKUs and the second fails, the first one stays assigned.
- The response surfaces business-logic errors as 400 with a specific error code in `error.code`, see section 8.

### 5.3 Remove license

| Property | Value |
|---|---|
| Endpoint | `POST /users/{id}/assignLicense` (same endpoint, different payload) |
| Permission | Same as 5.2 |
| Flow step | 7.6 `HTTP_RemoveLicenses` |
| Success status | 200 OK |

Request:

```json
{
  "addLicenses": [],
  "removeLicenses": [
    "c7df2760-2c81-4ef7-b578-5b5392b571df"
  ]
}
```

`removeLicenses` is an array of skuId GUIDs only (not objects). The flow builds this array with a Select action from `body('HTTP_GetTargetUser')?['assignedLicenses']` mapped to `item()?['skuId']`.

Quirks:

- Removing a license that the user does not have returns 400 with `Specified user does not have one or more of the licenses to be removed`. The flow only requests removal of currently assigned SKUs, so this should not happen.
- Licenses assigned via group-based licensing cannot be removed here, they return 400 with `License cannot be removed because user is assigned the license via group membership`. The termination flow handles this implicitly by removing the user from groups first (step 7.5) before this step.

---

## 6. Mailbox configuration

### 6.1 Set auto-reply via mailboxSettings

| Property | Value |
|---|---|
| Endpoint | `PATCH /users/{id}/mailboxSettings` |
| Permission | `MailboxSettings.ReadWrite` |
| Flow step | 7.7 `HTTP_SetForwarding` |
| Success status | 200 OK |

Request:

```json
{
  "automaticRepliesSetting": {
    "status": "alwaysEnabled",
    "internalReplyMessage": "I am no longer with the company. For assistance, contact delegate@yourtenant.onmicrosoft.com.",
    "externalReplyMessage": "I am no longer with the company. For assistance, contact delegate@yourtenant.onmicrosoft.com.",
    "externalAudience": "all"
  }
}
```

Quirks:

- `status` values: `disabled`, `alwaysEnabled`, `scheduled`. Use `scheduled` with `scheduledStartDateTime` and `scheduledEndDateTime` for a date range.
- `externalAudience` values: `none`, `contactsOnly`, `all`.
- The mailbox must exist. For a brand-new user, the mailbox is provisioned within ~5 minutes of license assignment but can take up to an hour. Calling this endpoint before the mailbox exists returns 404 with `MailboxNotEnabledForRESTAPI`. The flow is termination-only here, so the mailbox always exists by then.
- Setting auto-reply does not actually forward incoming mail to the delegate, it only sends an OOO reply. The inbox rule in 6.2 does the real forwarding.

### 6.2 Create inbox forwarding rule

| Property | Value |
|---|---|
| Endpoint | `POST /users/{id}/mailFolders/inbox/messageRules` |
| Permission | `Mail.ReadWrite` |
| Flow step | 7.7 `HTTP_CreateForwardRule` |
| Success status | 201 Created |

Request:

```json
{
  "displayName": "Forward to delegate after termination",
  "sequence": 1,
  "isEnabled": true,
  "actions": {
    "forwardTo": [
      {
        "emailAddress": {
          "address": "delegate@yourtenant.onmicrosoft.com",
          "name": "Delegate Person"
        }
      }
    ],
    "stopProcessingRules": false
  }
}
```

The `name` field inside `emailAddress` is optional. Graph resolves the display name from the address if omitted.

Quirks:

- `sequence` controls rule order. 1 means this rule runs first. If the mailbox has existing rules, choose a sequence that does not collide, or use `999` to run last.
- `stopProcessingRules: false` lets the message continue through other rules after this one. For an offboarding rule, you usually want this so journaling and retention rules still run.
- To also keep a copy in the inbox while forwarding, use `actions.forwardTo` with `stopProcessingRules: false`. To forward and delete, use `actions.delete: true` alongside `forwardTo`.
- Forwarding to an external (non-tenant) address may be blocked by an Exchange Online remote-domain or anti-spam outbound policy. Test the destination address first.
- An alternative is the legacy mailbox property `ForwardingSmtpAddress` set via Exchange Online PowerShell. The inbox-rule approach is preferred because it is visible to the user in Outlook rules and survives mailbox moves.

---

## 7. Endpoints used outside Graph (cross-reference)

The flow also calls non-Graph APIs. These are documented in the connectors' own pages, listed here so this file is a complete API map.

| Operation | API | Connector / mechanism |
|---|---|---|
| Send welcome email | Office 365 Outlook | "Send an email (V2)" action |
| Send termination notification | Office 365 Outlook | "Send an email (V2)" action |
| Convert mailbox to shared | Exchange Online PowerShell `Set-Mailbox -Type Shared` | Azure Automation runbook `Convert-MailboxToShared` invoked via the "Create job" action of the Azure Automation connector |
| Create / read / update audit row | SharePoint REST | SharePoint connector "Create item", "Get item", "Update item" |
| Read form response | Microsoft Forms | Forms connector "Get response details" |

---

## 8. Error responses

Graph returns errors in a standard envelope:

```json
{
  "error": {
    "code": "Request_BadRequest",
    "message": "Specified license is invalid.",
    "innerError": {
      "date": "2026-05-15T18:42:11",
      "request-id": "...",
      "client-request-id": "..."
    }
  }
}
```

The flow's error-handler block (Deliverable 2 section 4, block 3) captures the full response body into `varErrorDetails` for the audit row.

### Common errors per endpoint

| Status | Code | Where it appears | Cause and handling |
|---|---|---|---|
| 400 | `Request_BadRequest` | Create user | mailNickname has invalid chars, UPN domain not verified, passwordProfile fails policy. Flow logs and terminates as critical. |
| 400 | `UserNotLicensable` | Assign license | `usageLocation` not set on user. Flow's create-user step sets this, but if the user existed before the flow, you must PATCH `usageLocation` first. |
| 400 | `LicenseAssignmentAttachedToGroup` | Remove license | License is inherited from a group. Remove user from the group first (step 7.5 handles this). |
| 401 | `InvalidAuthenticationToken` | Any | Token expired or app secret rotated. Power Automate's HTTP action refreshes tokens automatically, so persistent 401 indicates a secret-expiry or permission issue. |
| 403 | `Authorization_RequestDenied` | Add/remove member | Group is on-prem synced, dynamic, or a system group. Flow logs and continues. |
| 403 | `Insufficient privileges` | Any | Admin consent not granted for the permission, or the permission was added but consent never re-applied. |
| 404 | `Request_ResourceNotFound` | Get user, set manager | UPN typo or user soft-deleted. Flow terminates as critical for target user, logs and continues for manager. |
| 404 | `MailboxNotEnabledForRESTAPI` | mailboxSettings or messageRules | Mailbox not yet provisioned. Add a Delay action or retry loop. Not expected in termination scenarios. |
| 409 | `Request_ConflictingObject` | Create user | UPN already exists. Choose a different UPN suffix or fail the flow. |
| 429 | (throttling) | Any high-volume endpoint | See section 10. Flow retries via the Do until 429 wrapper from Deliverable 2 section 8. |
| 500 / 502 / 503 / 504 | Transient | Any | Built-in exponential retry handles these. |

---

## 9. Pagination

Graph paginates with the OData `@odata.nextLink` property in the response. Default page size is 100, maximum 999 (via `$top`).

```json
{
  "value": [ /* ...100 items... */ ],
  "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=X..."
}
```

The flow paginates the one endpoint that can realistically exceed 100 results:

- `/users/{id}/memberOf` — implemented as a Do until loop in Deliverable 2 section 7.5 that follows `@odata.nextLink` until it returns null, concatenating each page's `value` array into a flow variable. This handles users with arbitrary numbers of group memberships.

Endpoints that do not need pagination handling in this flow:

- `/subscribedSkus` — capped at ~50 SKUs per tenant in practice, never paginates.
- `/groups?$filter=displayName eq ...` — a single display-name match returns one or zero results in almost all cases.

If you later add a step that calls `/users`, `/groups`, or `/auditLogs` without a tight filter, apply the same `@odata.nextLink` Do until pattern.

---

## 10. Throttling

Graph applies per-app and per-tenant throttling. Documented limits (subject to change, the canonical source is https://learn.microsoft.com/en-us/graph/throttling):

| Service | Limit |
|---|---|
| User write | 75 requests per app per tenant per 5 seconds |
| Group write | 75 requests per app per tenant per 5 seconds |
| Mail (sendMail, messageRules) | 4 requests per app per mailbox per second |
| Outlook REST (mailboxSettings) | 10,000 requests per mailbox per 10 minutes |

A throttled response is HTTP 429 with a `Retry-After` header in seconds. The flow's Do until wrapper (Deliverable 2 section 8) honors this header. Without the wrapper, the built-in HTTP retry policy uses exponential backoff and ignores the header, which works but may delay longer than necessary.

In a 25-user dev tenant with manual form submissions, throttling is unlikely. The wrapper exists to demonstrate the pattern and to handle the case where you submit 5-10 forms in rapid succession during testing.

### Throttling response example

```
HTTP/1.1 429 Too Many Requests
Retry-After: 17
Content-Type: application/json

{
  "error": {
    "code": "TooManyRequests",
    "message": "Application is over its first quota",
    "innerError": { ... }
  }
}
```

Wait 17 seconds before retrying. The flow's Do until reads `outputs('HTTP_Action')?['headers']?['Retry-After']` and delays accordingly.

---

## 11. Versioning and deprecation

This flow uses Graph v1.0 exclusively. None of the endpoints used here are scheduled for deprecation as of 2026-05-15.

Notable v1.0 endpoints in the broader Graph that this flow does not use:

- `/users/{id}/authentication/methods` (manage MFA methods) — would be useful for resetting a new hire's MFA but requires `UserAuthenticationMethod.ReadWrite.All`, an additional consent
- `/users/{id}/changePassword` — used only by the user themselves, not by admin apps
- `/identityProtection/riskyUsers` — useful for ad-hoc termination triggers based on risk

If you extend the flow later, prefer v1.0. Beta endpoints can change without notice, and they are not covered by Microsoft Graph SLAs.

---

## 12. Local testing with Graph Explorer

Before wiring an endpoint into Power Automate, test it manually in Graph Explorer at https://developer.microsoft.com/en-us/graph/graph-explorer. Sign in with your tenant admin account, consent to the same permissions the app registration has, and run each endpoint with a sample payload. This catches payload typos and permission gaps without burning Power Automate run-history entries.

For payloads with sensitive data (like the create-user passwordProfile), use throwaway test values, not real password material.

---

## Build checklist

- [ ] All 6 application permissions visible and admin-consented in the app registration
- [ ] Each endpoint in sections 2 through 6 verified working in Graph Explorer with the same admin account
- [ ] Sample SKU GUID retrieved from `/subscribedSkus` and recorded in the README
- [ ] At least one test user created and deleted via Graph Explorer using the section 2.1 payload
- [ ] A 429 simulated by burst-calling Graph Explorer (or accepted as untested, documented in the README test log)
- [ ] Power Automate HTTP action retry policy configured per Deliverable 2 section 4.1
