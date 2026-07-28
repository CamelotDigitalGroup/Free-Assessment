# 1. Create the app registration & certificate

This assessment authenticates to Microsoft 365 as an **app** (app-only), using a **certificate**
instead of a password. This page walks you through the one-time setup. It takes about 15 minutes
and a Global Administrator (or Privileged Role Administrator) is required to grant consent.

---

## 1. Register the application

1. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com).
2. Go to **Identity → Applications → App registrations → New registration**.
3. Name it something clear, e.g. **`M365 Security Assessment`**.
4. Under **Supported account types**, choose **Accounts in this organizational directory only**.
5. Leave Redirect URI blank and click **Register**.
6. On the app's **Overview** page, copy the **Application (client) ID** and **Directory (tenant) ID** —
   you'll pass these to the script as `-ClientId` and `-TenantId`.

---

## 2. Add the API permissions

### Microsoft Graph (application permissions)

1. Open your app → **API permissions → Add a permission → Microsoft Graph → Application permissions**.
2. Add each of the following (all are **read-only**):

   - `AuditLog.Read.All`
   - `Directory.Read.All`
   - `Policy.Read.All`
   - `DeviceManagementConfiguration.Read.All`
   - `DeviceManagementManagedDevices.Read.All`
   - `SecurityEvents.Read.All`
   - `Sites.Read.All`
   - `SharePointTenantSettings.Read.All` *(optional)*
   - `Reports.Read.All`
   - `RoleManagement.Read.Directory`
   - `IdentityRiskyUser.Read.All`
   - `SecurityActions.Read.All`

### Office 365 Exchange Online (application permission)

3. **Add a permission → APIs my organization uses →** search **Office 365 Exchange Online →
   Application permissions →** add **`Exchange.ManageAsApp`**.

### Grant consent

4. Click **Grant admin consent for &lt;your tenant&gt;** and confirm. All permissions should show a green tick.

---

## 3. Create the certificate

A self-signed certificate is fine. Run this in an **elevated PowerShell** on the machine that will run
the assessment:

```powershell
$cert = New-SelfSignedCertificate `
    -Subject "CN=M365SecurityAssessment" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -NotAfter (Get-Date).AddYears(1)

# Note the thumbprint — you'll pass it to the script as -CertThumbprint
$cert.Thumbprint

# Export the PUBLIC key (.cer) to upload to the app registration
Export-Certificate -Cert $cert -FilePath "$HOME\M365SecurityAssessment.cer"
```

> The **private key stays on your machine** in your certificate store and is never uploaded anywhere.
> Only the public `.cer` is uploaded to Entra.

---

## 4. Upload the certificate to the app

1. In your app → **Certificates & secrets → Certificates → Upload certificate**.
2. Select the `M365SecurityAssessment.cer` file you exported and click **Add**.
3. Confirm the thumbprint shown matches the one printed in step 3.

---

## 5. Assign a directory role (for audit-log / Purview checks)

1. Go to **Identity → Roles & admins → Roles & admins**.
2. Open **Global Reader** *(or **Compliance Reader**)*.
3. **Add assignments →** search for your app by name → add it.

---

## You're ready

You now have the three values the script needs:

| Value | Where it came from |
| ----- | ------------------ |
| `-TenantId` | App **Overview** → Directory (tenant) ID |
| `-ClientId` | App **Overview** → Application (client) ID |
| `-CertThumbprint` | Printed in step 3 (also shown under Certificates & secrets) |

Continue to **[02-RUNNING-THE-ASSESSMENT.md](02-RUNNING-THE-ASSESSMENT.md)**.
