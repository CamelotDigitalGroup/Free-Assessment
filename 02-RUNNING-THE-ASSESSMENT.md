# 2. Running the assessment

Once the [app registration & certificate](01-APP-REGISTRATION.md) are in place, running the
assessment is two commands.

---

## 1. Install the optional modules (recommended)

```powershell
# Exchange Online domain (if omitted, Exchange checks are skipped)
Install-Module ExchangeOnlineManagement -Scope CurrentUser

# Excel evidence workbook (if omitted, raw data falls back to CSV)
Install-Module ImportExcel -Scope CurrentUser
```

For the PDF report, make sure **Microsoft Edge** or **Google Chrome** is installed (both ship with
Windows / are easy to install). If neither is present, you still get the HTML report.

---

## 2. Run the assessment (read-only)

```powershell
.\M365-SecurityAssessment.ps1 `
    -TenantId       "f7c53549-c546-4fab-bab5-8d45ce67f292" `
    -ClientId       "274c5433-710f-46f1-b303-97ec0f5255b8" `
    -CertThumbprint "B61EDAABAF34649F69959391E1B02D11E7CC1ACE" `
    -OutputFolder   "C:\Assessments\Output"
```

**Optional tuning parameters:**

| Parameter | Default | Meaning |
| --------- | ------- | ------- |
| `-MinutesBack` | `1440` | Look-back window (minutes) for audit / sign-in log queries |
| `-DormantDays` | `45` | Days of inactivity before an account is flagged dormant |
| `-StaleDeviceDays` | `30` | Days since last check-in before a device is flagged stale |

The script prints a live **PASS / FAIL / WARN / MANUAL** summary per domain and writes to
`-OutputFolder`:

- `Assessment_Results_<timestamp>.json`
- `Assessment_Workbook_<timestamp>.xlsx` *(or CSVs if ImportExcel isn't installed)*

---

## 3. Build the report

Point the report script at the JSON the assessment just produced:

```powershell
.\New-M365Report.ps1 `
    -JsonPath     "C:\Assessments\Output\Assessment_Results_20260728_101500.json" `
    -CustomerName "Your Company Ltd" `
    -OutputFolder "C:\Assessments\Output"
```

**Optional parameters:**

| Parameter | Default | Meaning |
| --------- | ------- | ------- |
| `-CustomerName` | auto-detected | Name shown on the cover page |
| `-LogoPath` | built-in shield | Custom PNG/SVG logo for the cover |
| `-PreparedBy` | `Camelot Digital Group` | Preparer name |
| `-TypicalClientScore` | `85` | Benchmark score for the comparison bar |

You'll get:

- `Assessment_Report_<timestamp>.html`
- `Assessment_Report_<timestamp>.pdf`

---

## Troubleshooting

| Symptom | Fix |
| ------- | --- |
| `Certificate with thumbprint '…' not found` | The certificate isn't in `Cert:\CurrentUser\My` or `Cert:\LocalMachine\My` on this machine. Re-run step 3 of the app-registration guide on the machine you're running from. |
| Exchange checks show `MANUAL — module not installed` | `Install-Module ExchangeOnlineManagement`. |
| Workbook is CSVs, not `.xlsx` | `Install-Module ImportExcel`. |
| No PDF, only HTML | Install Microsoft Edge or Google Chrome, then re-run the report script. |
| `403 / insufficient privileges` on some checks | Confirm **admin consent** was granted and the app has the **Global Reader / Compliance Reader** role assigned. |
| Some controls always `MANUAL` | A few controls require human judgement and are intentionally left for manual review. |
