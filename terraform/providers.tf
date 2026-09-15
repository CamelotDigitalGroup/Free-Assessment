terraform {
  required_version = ">= 1.5.0"

  required_providers {
    azuread = {
      source  = "hashicorp/azuread"
      version = "~> 3.0"
    }
  }
}

# Authenticates using whatever the Azure CLI / environment is signed in as.
# The identity running Terraform must be a Global Administrator or Privileged
# Role Administrator so it can grant admin consent and assign the directory role.
provider "azuread" {
  # tenant_id is picked up automatically from `az login`.
  # Override explicitly if you manage multiple tenants:
  # tenant_id = var.tenant_id
}