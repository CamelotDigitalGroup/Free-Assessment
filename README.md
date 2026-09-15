<div align="center">

# Camelot Digital Group — Free Microsoft 365 Security Assessment

**A read-only, certificate-authenticated security assessment for Microsoft 365 tenants — fully open source, so you can read every line before you run it.**

[![PowerShell](https://img.shields.io/badge/PowerShell-7.2%2B-5391FE?logo=powershell&logoColor=white)](https://learn.microsoft.com/powershell/)
[![Microsoft 365](https://img.shields.io/badge/Microsoft%20365-Graph%20%2B%20Exchange%20Online-D83B01?logo=microsoft&logoColor=white)](https://learn.microsoft.com/graph/)
[![Auth](https://img.shields.io/badge/Auth-Certificate%20(app--only)-1B4332)](01-APP-REGISTRATION.md#3-create-the-certificate)
[![Access](https://img.shields.io/badge/Access-Read--only-22c55e)](#security--privacy)
[![License: MIT](https://img.shields.io/badge/License-MIT-D4A843)](LICENSE)

</div>

---

## What this is

This repository contains the two PowerShell scripts Camelot Digital Group uses to run its
**Advanced** Microsoft 365 security assessment — the supervised, certificate-based review that
goes deepest, including a full **Exchange Online** analysis across all seven Microsoft 365
security domains.

| Script | What it does |
| ------ | ------------ |
| **`M365-SecurityAssessment.ps1`** | Connects to your tenant **read-only** using certificate-based app authentication, evaluates 65 security controls (71 individual checks) across Entra ID, Exchange, Teams, Intune, SharePoint/OneDrive, Defender and Purview, and writes a structured JSON results file plus an Excel workbook of the raw evidence. |
| **`New-M365Report.ps1`** | Turns that JSON into a branded, self-contained **HTML report** and renders it to **PDF** — clear findings, severities and recommendations, ready to hand to leadership. |

Everything is **read-only**. The scripts can never change, delete, move or send anything in
your environment, and they never read the contents of your emails or files.

---

## The seven domains assessed

1. **Entra ID** — identity, MFA coverage, admin accounts, Conditional Access, legacy auth, PIM
2. **Exchange Online** — mailbox auditing, transport rules, forwarding, authentication
3. **Microsoft Teams** — external / guest access and meeting policies
4. **Intune** — device compliance, configuration and app protection policies
5. **SharePoint & OneDrive** — tenant and per-site external sharing
6. **Microsoft Defender** — Secure Score, alerts, risky users
7. **Purview** — audit log and compliance posture

Each control is scored **PASS / FAIL / WARN / MANUAL** with a severity rating.

---

## Quick start

> Full, click-by-click setup is in **[01-APP-REGISTRATION.md](01-APP-REGISTRATION.md)**
> and **[02-RUNNING-THE-ASSESSMENT.md](02-RUNNING-THE-ASSESSMENT.md)**.
>
> Prefer infrastructure-as-code? The whole Entra setup (app registration, read-only
> permissions, admin consent, certificate upload and directory role) can be provisioned
> automatically with the **[Terraform module in `terraform/`](terraform/README.md)**.

```powershell
# 1) Run the assessment (read-only) against your tenant
.\M365-SecurityAssessment.ps1 `
    -TenantId       "<your-tenant-guid>" `
    -ClientId       "<app-registration-client-id>" `
    -CertThumbprint "<certificate-thumbprint>" `
    -OutputFolder   "C:\Assessments\Output"

# 2) Build the branded PDF report from the JSON it produced
.\New-M365Report.ps1 `
    -JsonPath     "C:\Assessments\Output\Assessment_Results_<timestamp>.json" `
    -CustomerName "Your Company Ltd" `
    -OutputFolder "C:\Assessments\Output"
```

---

## Prerequisites

| Requirement | Notes |
| ----------- | ----- |
| **PowerShell 7.2+** | Recommended on Windows. `pwsh` cross-platform also works for the Graph portion. |
| **An Entra app registration** | App-only, with a certificate — set up manually via [01-APP-REGISTRATION.md](01-APP-REGISTRATION.md), or automatically with the [Terraform module](terraform/README.md). |
| **A certificate** | Self-signed is fine. Public key uploaded to the app; private key installed in your local certificate store. |
| **`ExchangeOnlineManagement` module** | For the Exchange Online domain. If missing, Exchange checks are skipped gracefully. `Install-Module ExchangeOnlineManagement` |
| **`ImportExcel` module** *(optional)* | For the `.xlsx` evidence workbook. If missing, raw data falls back to CSV. `Install-Module ImportExcel` |
| **Microsoft Edge or Google Chrome** | Used by the report script (headless) to render the PDF. If neither is present, the HTML report is still produced. |

---

## Required app permissions

**Microsoft Graph — Application permissions** (admin consent required):

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

**Exchange Online — Application permission:**

| Permission | Used for |
| ---------- | -------- |
| `Exchange.ManageAsApp` | App-only Exchange Online cmdlets (`Get-*`, `Search-UnifiedAuditLog`) |

**Entra directory role** (assign to the app's service principal): **Global Reader** *or* **Compliance Reader**
(needed for the unified audit log / Purview checks).

> Every permission above is a **`.Read.` / read-only** permission. There are no write, send or delete scopes.

---

## What you get

- **`Assessment_Results_<timestamp>.json`** — every control, status, severity and evidence detail
- **`Assessment_Workbook_<timestamp>.xlsx`** — raw data tabs (Users, Devices, Policies, Risky Users, …)
- **`Assessment_Report_<timestamp>.html`** — branded, self-contained report (no external dependencies)
- **`Assessment_Report_<timestamp>.pdf`** — the same report as a shareable PDF

---

## Security & privacy

- **Read-only, always.** Only `*.Read.*` Graph scopes and read-only Exchange cmdlets are used — the
  scripts cannot change, delete, move or send anything, and never read the content of your mail or files.
- **Certificate-based, app-only.** No passwords are stored or transmitted. Authentication uses a signed
  JWT assertion against your certificate's private key, which never leaves your machine.
- **You stay in control.** Remove the app registration (or delete the certificate) at any time to revoke
  all access instantly.
- **Fully open source.** Read every line in this repository before you run a thing.

---

## Prefer not to run it yourself?

Camelot Digital Group offers this as a **supervised, white-glove assessment**: we run the entire process
live with you on a screen-share so you can watch every step, then remove all access and delete the
certificate the moment we're done.

**[Request an assessment →](https://camelotdigitalgroup.com/free-assessment)**

---

## License

Released under the [MIT License](LICENSE). © Camelot Digital Group.
