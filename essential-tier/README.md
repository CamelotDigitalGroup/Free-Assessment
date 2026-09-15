# Essential tier — the script behind the free self-serve assessment

This is the exact script that runs when you request the **free Essential assessment** at
[camelotdigitalgroup.com/free-assessment](https://camelotdigitalgroup.com/free-assessment). It's
published here for the same reason as the [Advanced-tier script](../M365-SecurityAssessment.ps1)
one level up: so you can read every line before — or after — granting access.

## This is not something you run yourself

Unlike the Advanced-tier script, you don't download this, generate a certificate, or run
PowerShell against your own tenant. The flow is:

1. You click **Request assessment** on the website and sign in with your Microsoft 365 admin
   account.
2. Microsoft's own standard admin-consent screen shows you exactly what Camelot's app is asking
   for (the permission list below) — you approve or decline there, nothing happens without it.
3. Camelot's Azure Automation account runs **this exact file** against your tenant, using a
   client secret Camelot holds for its own shared, multi-tenant app registration — not a
   credential you create or manage.
4. The report is emailed to you automatically. No certificate, no local setup, no PowerShell.

That's the whole point of the free tier: zero setup on your end. This file being public is what
lets you verify that claim rather than take our word for it.

## Why this tier is narrower than Advanced

This script authenticates to **Microsoft Graph only**, via a plain OAuth client-secret
(client-credentials flow) — no certificate, and critically, **no Exchange Online or Microsoft
Teams PowerShell connection at all**. Both of those require certificate-based app-only auth,
which this tier deliberately doesn't use. Any control that depends on Exchange or Teams data
(mailbox/transport rules, Teams external-access policies, etc.) is reported as requiring manual
review rather than silently skipped or guessed at — see the [Advanced tier](../) if you want
those checked automatically too, under supervision, via a screen-share.

## Required app permissions (Microsoft Graph only)

| Permission | Used for |
| ---------- | -------- |
| `AuditLog.Read.All` | Sign-in + Entra audit logs |
| `Directory.Read.All` | Users, groups, devices, apps, roles |
| `Policy.Read.All` | Conditional Access & authentication policies |
| `DeviceManagementConfiguration.Read.All` | Intune configuration policies |
| `DeviceManagementManagedDevices.Read.All` | Intune device inventory |
| `SecurityEvents.Read.All` | Defender alerts / Secure Score |
| `Sites.Read.All` | SharePoint site enumeration & sharing links |
| `SharePointTenantSettings.Read.All` | Tenant-wide sharing settings *(optional)* |
| `Reports.Read.All` | Usage / MFA registration reports |
| `RoleManagement.Read.Directory` | Role assignments (PIM check) |
| `IdentityRiskyUser.Read.All` | Risky user list |
| `SecurityActions.Read.All` | Defender for Cloud Apps |

Every permission is a **`.Read.` / read-only** Microsoft Graph scope. No Exchange Online or Teams
permission is requested for this tier, and there are no write, send or delete scopes anywhere in
the list.

## What you get

The same [`New-M365Report.ps1`](../New-M365Report.ps1) used by the Advanced tier turns the JSON
this script produces into the branded PDF report you receive by email — it's genuinely the same
renderer for both tiers, just fed a narrower set of results.
