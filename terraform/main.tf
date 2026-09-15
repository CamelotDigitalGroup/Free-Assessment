###############################################################################
# Well-known service principals (Microsoft Graph + Exchange Online)
#
# We look up the app-role IDs by name at plan time, so there are no fragile
# hardcoded GUIDs. `use_existing = true` reuses the SPs already present in the
# tenant instead of trying to create them.
###############################################################################

data "azuread_client_config" "current" {}

data "azuread_application_published_app_ids" "well_known" {}

resource "azuread_service_principal" "msgraph" {
  client_id    = data.azuread_application_published_app_ids.well_known.result["MicrosoftGraph"]
  use_existing = true
}

resource "azuread_service_principal" "exo" {
  # Office 365 Exchange Online
  client_id    = "00000002-0000-0ff1-ce00-000000000000"
  use_existing = true
}

locals {
  # Read-only Microsoft Graph application permissions required by the assessment.
  graph_roles = compact([
    "AuditLog.Read.All",
    "Directory.Read.All",
    "Policy.Read.All",
    "DeviceManagementConfiguration.Read.All",
    "DeviceManagementManagedDevices.Read.All",
    "SecurityEvents.Read.All",
    "Sites.Read.All",
    var.include_optional_sharepoint_setting ? "SharePointTenantSettings.Read.All" : "",
    "Reports.Read.All",
    "RoleManagement.Read.Directory",
    "IdentityRiskyUser.Read.All",
    "SecurityActions.Read.All",
  ])
}

###############################################################################
# The app registration
###############################################################################

resource "azuread_application" "assessment" {
  display_name     = var.application_name
  sign_in_audience = "AzureADMyOrg"

  # Microsoft Graph (application/read-only)
  required_resource_access {
    resource_app_id = data.azuread_application_published_app_ids.well_known.result["MicrosoftGraph"]

    dynamic "resource_access" {
      for_each = toset(local.graph_roles)
      content {
        id   = azuread_service_principal.msgraph.app_role_ids[resource_access.value]
        type = "Role"
      }
    }
  }

  # Office 365 Exchange Online (application/read-only cmdlets)
  required_resource_access {
    resource_app_id = "00000002-0000-0ff1-ce00-000000000000"

    resource_access {
      id   = azuread_service_principal.exo.app_role_ids["Exchange.ManageAsApp"]
      type = "Role"
    }
  }
}

resource "azuread_service_principal" "assessment" {
  client_id = azuread_application.assessment.client_id
}

###############################################################################
# Certificate (public key) upload
###############################################################################

resource "azuread_application_certificate" "assessment" {
  count          = var.certificate_path == null ? 0 : 1
  application_id = azuread_application.assessment.id
  type           = "AsymmetricX509Cert"
  value          = file(var.certificate_path)
  encoding       = "pem"
}

###############################################################################
# Admin consent — grant every requested app role to the assessment SP
###############################################################################

resource "azuread_app_role_assignment" "graph" {
  for_each            = toset(local.graph_roles)
  app_role_id         = azuread_service_principal.msgraph.app_role_ids[each.value]
  principal_object_id = azuread_service_principal.assessment.object_id
  resource_object_id  = azuread_service_principal.msgraph.object_id
}

resource "azuread_app_role_assignment" "exo" {
  app_role_id         = azuread_service_principal.exo.app_role_ids["Exchange.ManageAsApp"]
  principal_object_id = azuread_service_principal.assessment.object_id
  resource_object_id  = azuread_service_principal.exo.object_id
}

###############################################################################
# Directory role (Global Reader / Compliance Reader) for audit-log & Purview
###############################################################################

resource "azuread_directory_role" "assessment" {
  display_name = var.directory_role
}

resource "azuread_directory_role_assignment" "assessment" {
  role_id             = azuread_directory_role.assessment.template_id
  principal_object_id = azuread_service_principal.assessment.object_id
}