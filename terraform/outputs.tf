output "tenant_id" {
  description = "Pass to the script as -TenantId"
  value       = data.azuread_client_config.current.tenant_id
}

output "client_id" {
  description = "Pass to the script as -ClientId"
  value       = azuread_application.assessment.client_id
}

output "service_principal_object_id" {
  description = "Object ID of the app's service principal (useful for auditing role assignments)."
  value       = azuread_service_principal.assessment.object_id
}

output "next_steps" {
  value = <<-EOT
    App registration created and consented.

    1) Make sure your certificate's PRIVATE key is installed in the local
       certificate store on the machine that will run the assessment.
    2) Run:
         .\M365-SecurityAssessment.ps1 \n             -TenantId       "${data.azuread_client_config.current.tenant_id}" \n             -ClientId       "${azuread_application.assessment.client_id}" \n             -CertThumbprint "<your-cert-thumbprint>" \n             -OutputFolder   "C:\Assessments\Output"
  EOT
}