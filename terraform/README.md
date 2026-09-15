# Terraform — automated Entra setup

This module provisions **everything the assessment scripts need** in Microsoft Entra
(Azure AD), so you don't have to click through the portal:

- the **app registration** (app-only, single-tenant)
- a **service principal** for it
- all the **read-only Microsoft Graph** application permissions
- the **Exchange Online** `Exchange.ManageAsApp` permission
- **admin consent** for every one of those permissions
- the **Global Reader** (or Compliance Reader) directory role assignment
- *(optional)* uploads your **public certificate** to the app

The app-role IDs are looked up by name at plan time, so there are no hardcoded GUIDs.

---

## Prerequisites

| Requirement | Notes |
| ----------- | ----- |
| **Terraform ≥ 1.5** | <https://developer.hashicorp.com/terraform/install> |
| **Azure CLI** | Sign in with `az login` as a **Global Administrator** or **Privileged Role Administrator** (needed to grant consent and assign the directory role). |
| **A certificate** *(optional here)* | You can let Terraform upload the public `.cer`, or add it later in the portal. The **private key** must live in the certificate store of the machine that runs the assessment. See [../docs/01-APP-REGISTRATION.md](../docs/01-APP-REGISTRATION.md). |

---

## Usage

```bash
# 1) Sign in to the tenant you want to assess
az login --tenant <your-tenant-guid>

# 2) Configure your variables
cp terraform.tfvars.example terraform.tfvars
#   edit terraform.tfvars (app name, certificate_path, directory_role, ...)

# 3) Provision
terraform init
terraform plan
terraform apply
```

When it finishes, Terraform prints the `tenant_id` and `client_id` to pass straight
into the script:

```powershell
.\M365-SecurityAssessment.ps1 `
    -TenantId       "<output.tenant_id>" `
    -ClientId       "<output.client_id>" `
    -CertThumbprint "<your-cert-thumbprint>" `
    -OutputFolder   "C:\Assessments\Output"
```

---

## Generating the certificate

If you want Terraform to upload the certificate, first create it (see
[../docs/01-APP-REGISTRATION.md](../docs/01-APP-REGISTRATION.md) for detail):

```powershell
$cert = New-SelfSignedCertificate `
    -Subject "CN=M365SecurityAssessment" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable -KeySpec Signature `
    -KeyLength 2048 -NotAfter (Get-Date).AddYears(1)

$cert.Thumbprint                                   # -> -CertThumbprint
Export-Certificate -Cert $cert -FilePath ".\M365SecurityAssessment.cer"
```

Then set `certificate_path = "./M365SecurityAssessment.cer"` in `terraform.tfvars`.

> Terraform only ever handles the **public** key. The private key never leaves your machine.

---

## Cleaning up

To remove the app registration and all access it was granted:

```bash
terraform destroy
```

---

## Notes

- Directory-role changes and admin consent can take a minute or two to propagate before
  the assessment will authenticate successfully.
- State files (`*.tfstate`) and `terraform.tfvars` can contain identifiers for your tenant
  — they're git-ignored here. Keep them private.