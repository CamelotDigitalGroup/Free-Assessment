<div align="center">

# Camelot Digital Group — Free Microsoft 365 Security Assessment

**The exact script that runs when you request our free assessment — published so you can read every line before you grant a thing.**

[![PowerShell](https://img.shields.io/badge/PowerShell-7.2%2B-5391FE?logo=powershell&logoColor=white)](https://learn.microsoft.com/powershell/)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-Read--only-D83B01?logo=microsoft&logoColor=white)](https://learn.microsoft.com/graph/)
[![Auth](https://img.shields.io/badge/Auth-OAuth%20client--secret-1B4332)](#how-it-actually-runs)
[![Access](https://img.shields.io/badge/Access-Read--only-22c55e)](#security--privacy)
[![License: MIT](https://img.shields.io/badge/License-MIT-D4A843)](LICENSE)

</div>

---

## What this is

This is the real, unmodified script Camelot Digital Group's own backend runs the moment you
request a free assessment at
**[camelotdigitalgroup.com/free-assessment](https://camelotdigitalgroup.com/free-assessment)**
and grant consent. Not a sanitized example — the same file, served to our automation at request
time from this exact source.

| Script | What it does |
| ------ | ------------ |
| **`M365-SecurityAssessment.ps1`** | Connects to your tenant **read-only** via a plain OAuth client-secret (Microsoft Graph only), evaluates security controls across Entra ID, Intune, SharePoint/OneDrive, Defender and Purview, and writes a structured JSON results file plus an Excel workbook of the raw evidence. |
| **`New-M365Report.ps1`** | Turns that JSON into a branded, self-contained **HTML report** and renders it to **PDF** — clear findings, severities and recommendations. |

Everything is **read-only**. The script can never change, delete, move or send anything in your
environment, and never reads the content of your emails or files.

---

## How it actually runs

You never touch PowerShell or set up an app registration yourself. The flow is:

1. You request an assessment on the website and sign in with your Microsoft 365 admin account.
2. Microsoft's own standard admin-consent screen shows you exactly what Camelot's app is asking
   for (the permission list below) — you approve or decline there.
3. Camelot's Azure Automation account runs **this exact file** against your tenant, using a
   client secret Camelot holds for its own shared, multi-tenant app registration — not a
   credential you create or manage.
4. The report is emailed to you automatically. No certificate, no local setup.

This file being public is what lets you verify that story rather than take our word for it.

---

## Scope

This assessment authenticates to **Microsoft Graph only** — no Exchange Online or Microsoft
Teams PowerShell connections, since those require certificate-based app-only auth, which this
free tier deliberately doesn't use. Any control that depends on Exchange or Teams data
(mailbox/transport rules, Teams external-access policies, etc.) is reported as requiring manual
review, not silently skipped or guessed at.

Want deeper coverage, including a full Exchange Online and Teams review? That's our **Advanced**
assessment — a supervised, certificate-based engagement we run live with you on a call, then
fully clean up afterwards. It isn't published in this repo (elevated, standing access shouldn't
be something anyone can run unsupervised), but if you're working with us on one, we'll share that
script with you directly as part of the engagement. Nothing about it is a black box — it's
simply not broadcast to the public internet by default. **[Get in touch](https://camelotdigitalgroup.com/free-assessment)**
to arrange one.

---

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
permission is requested, and there are no write, send or delete scopes anywhere in the list.

---

## What you get

- **`Assessment_Results_<timestamp>.json`** — every control, status, severity and evidence detail
- **`Assessment_Workbook_<timestamp>.xlsx`** — raw data tabs (Users, Devices, Policies, …)
- **`Assessment_Report_<timestamp>.html`** — branded, self-contained report
- **`Assessment_Report_<timestamp>.pdf`** — the same report as a shareable PDF, emailed to you

---

## Security & privacy

- **Read-only, always.** Only `*.Read.*` Graph scopes are used — the script cannot change,
  delete, move or send anything, and never reads the content of your mail or files.
- **You stay in control.** Remove the app registration from Enterprise Applications at any time
  to revoke all access instantly — we'll show you exactly how, right after your report arrives.
- **Fully open source.** Read every line in this repository before you grant a thing.

---

## License

Released under the [MIT License](LICENSE). © Camelot Digital Group.
