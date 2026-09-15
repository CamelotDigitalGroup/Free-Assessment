variable "tenant_id" {
  type        = string
  description = "The Entra (Azure AD) tenant GUID the app will be registered in. Optional if your `az login` session already targets the correct tenant."
  default     = null
}

variable "application_name" {
  type        = string
  description = "Display name for the app registration."
  default     = "M365 Security Assessment"
}

variable "certificate_path" {
  type        = string
  description = "Path to the PUBLIC certificate (.cer/.pem, base64/PEM encoded) to upload to the app registration. Generate it with New-SelfSignedCertificate + Export-Certificate (see docs/01-APP-REGISTRATION.md). Leave null to add the certificate manually later in the portal."
  default     = null
}

variable "directory_role" {
  type        = string
  description = "Directory role to assign to the app's service principal for audit-log / Purview checks. Use 'Global Reader' or 'Compliance Reader'."
  default     = "Global Reader"
}

variable "include_optional_sharepoint_setting" {
  type        = bool
  description = "Also grant the optional SharePointTenantSettings.Read.All permission."
  default     = true
}