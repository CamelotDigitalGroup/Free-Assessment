<#
.SYNOPSIS
    M365 Cloud Security Assessment Script - ESSENTIAL TIER (Graph-only,
    client-secret). Run automatically by the Azure Automation runbook
    behind camelotdigitalgroup.com's self-serve free-assessment funnel -
    served to that runbook at request time from this exact file (see
    app/api/assessment/runner-scripts/[name]/route.ts), not a build-time
    snapshot, so this file is always what actually runs.

    NOT the certificate-based Advanced-tier script Camelot staff run
    manually for a supervised assessment, or publish to the public GitHub
    repo - that's a separate file at free-assessment/M365-SecurityAssessment.ps1
    in this same repo, with real Exchange/Teams PowerShell connectivity.
    Both happen to share this filename in their own directory; don't
    assume a change in one needs making in the other. The report renderer
    (New-M365Report.ps1) IS shared between them (copied in both places,
    kept in sync by hand) since JSON->HTML rendering is genuinely
    tier-agnostic.

    Reverse-engineered to reproduce the Sourcepass Cloud Assessment Report
    covering all 7 domains: Entra ID, Exchange, Teams, Intune,
    SharePoint/OneDrive, Defender, and Purview.

.DESCRIPTION
    Uses client-credentials app-only authentication (Microsoft Graph API
    only - no certificate, no Exchange Online/Teams PowerShell) to query
    tenant security configuration and evaluate each control against the
    same benchmark used in the report.

    Output:
        1. A structured JSON results file (one object per control)
        2. A companion Excel/CSV workbook with raw data tabs matching
           the reference workbook (Users, Devices, Policies, etc.)
        3. Console summary with Pass / Fail / Manual per domain

.REQUIRED APP PERMISSIONS
    Microsoft Graph (application):
        AuditLog.Read.All              sign-in + Entra audit logs
        Directory.Read.All             user/group/device/app/role reads
        Policy.Read.All                Conditional Access, auth policies
        DeviceManagementConfiguration.Read.All   Intune config policies
        DeviceManagementManagedDevices.Read.All  Intune device inventory
        SecurityEvents.Read.All        Defender alerts / secure score
        Sites.Read.All                 SharePoint full-tenant site enumeration
                                       (/sites/getAllSites) + per-site sharing links
        SharePointTenantSettings.Read.All  SharePoint tenant-wide sharing settings
                                       (/admin/sharepoint/settings) - OPTIONAL;
                                       if missing, sharingCapability reports 'unknown'
        Reports.Read.All               usage / MFA registration reports
        RoleManagement.Read.Directory  role assignments (PIM check)
        IdentityRiskyUser.Read.All     risky user list
        SecurityActions.Read.All       Defender for Cloud Apps

    That's the complete list. This is a Graph-only, client-secret build -
    $SkipExchange is hardcoded true (see below), so no Exchange Online or
    Teams PowerShell connection is ever made, and no directory role is
    needed. An earlier version of this comment listed Exchange.ManageAsApp
    and a Global Reader/Compliance Reader role requirement left over from
    when this was cloned from the Advanced-tier (certificate-based) script
    - confirmed those were dead requirements: Connect-ExchangeOnline only
    ever runs inside `if (-not $SkipExchange)`, which never evaluates true
    here.

.PARAMETER TenantId
    Azure AD / Entra tenant GUID.

.PARAMETER ClientId
    App registration (service principal) client ID.

.PARAMETER ClientSecret
    Client secret value from the app registration (Certificates & secrets ->
    Client secrets). Used for the Graph client-credentials flow. No certificate
    is required in this Graph-only build.

.PARAMETER OutputFolder
    Path where JSON results, CSV tabs, and the summary XLSX are written.
    Created if it does not exist.

.PARAMETER MinutesBack
    Look-back window for audit / sign-in log queries (default 1440 = 24 h).

.PARAMETER DormantDays
    Days of sign-in inactivity before an account is considered dormant
    (default 45, matching CIS benchmark).

.PARAMETER StaleDeviceDays
    Days since last check-in before a device is considered stale
    (default 30).

.EXAMPLE
    .\M365-SecurityAssessment.ps1 `
        -TenantId      "f7c53549-c546-4fab-bab5-8d45ce67f292" `
        -ClientId      "274c5433-710f-46f1-b303-97ec0f5255b8" `
        -ClientSecret  "<your-client-secret>" `
        -OutputFolder  "C:\Assessments\Output"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $TenantId,
    [Parameter(Mandatory)][string] $ClientId,
    [Parameter(Mandatory)][string] $ClientSecret,

    [string] $OutputFolder    = "C:\Assessments\Output",
    [int]    $MinutesBack     = 1440,
    [int]    $DormantDays     = 45,
    [int]    $StaleDeviceDays = 30
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region -- Helpers ------------------------------------------------------------

function Write-Section { param([string]$Title)
    Write-Host "`n$('-' * 70)" -ForegroundColor DarkGray
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "$('-' * 70)" -ForegroundColor DarkGray
}

function Write-Check { param([string]$Name,[string]$Status,[string]$Detail='')
    $colour = switch ($Status) {
        'PASS'   {'Green'}  'FAIL' {'Red'}
        'WARN'   {'Yellow'} default {'Gray'}
    }
    Write-Host ("  [{0,-6}] {1}" -f $Status, $Name) -ForegroundColor $colour
    if ($Detail) { Write-Host ("           {0}" -f $Detail) -ForegroundColor DarkGray }
}

# Build a result object stored in $Results
$Results = [System.Collections.Generic.List[PSCustomObject]]::new()
function Add-Result {
    param([string]$Domain,[string]$Control,[string]$Check,
          [string]$Status,[string]$Severity='medium',[string]$Detail='')
    $Results.Add([PSCustomObject]@{
        # Field names MUST match what New-M365Report.ps1 consumes
        # (Section / CheckId / Description), otherwise the report cannot
        # associate controls with their section pages and silently drops them.
        Section     = $Domain     # e.g. '1-EntraID'
        CheckId     = $Control    # e.g. '1.1'
        Description = $Check       # human-readable check text
        Status      = $Status      # PASS | FAIL | WARN | MANUAL
        Severity    = $Severity    # critical | high | medium | low
        Detail      = $Detail
        # Back-compat aliases (CSV / older consumers)
        Domain      = $Domain
        Control     = $Control
        Check       = $Check
    })
}

# Invoke Graph with auto-paging
# NOTE: strict mode throws on missing properties, so use PSObject.Properties to safely
# check for @odata.nextLink rather than accessing it directly.
function Invoke-GraphAll {
    param([string]$Uri,[hashtable]$Headers)
    $all  = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    do {
        $resp = Invoke-RestMethod -Uri $next -Headers $Headers -Method GET
        if ($null -ne $resp.value) { $all.AddRange([object[]]$resp.value) }
        # Safe property access - avoids strict-mode crash when nextLink is absent
        $nextProp = $resp.PSObject.Properties['@odata.nextLink']
        $next     = if ($nextProp) { $nextProp.Value } else { $null }
    } while ($next)
    # Unary comma is required: PowerShell unwraps a single-element collection
    # to its bare scalar element when it crosses a function return/pipeline
    # boundary, silently discarding .Count. Confirmed live against a real
    # tenant - this is a production bug affecting the automated pipeline,
    # not just the manual Advanced-tier script (same root cause found and
    # fixed there first, see free-assessment/M365-SecurityAssessment.ps1).
    # Any real customer tenant with exactly one matching Intune managed
    # device / config profile / dynamic group etc. would silently lose that
    # section's data under Set-StrictMode -Version Latest (above).
    return ,$all
}

# Strict-mode-safe nested property accessor.
# Returns $null (never throws) when any step in the dot-separated path is absent or null.
# Usage: Get-Prop $policy 'grantControls.builtInControls'
function Get-Prop {
    param($Obj, [string]$Path)
    $curr = $Obj
    foreach ($step in ($Path -split '\.')) {
        if ($null -eq $curr) { return $null }
        try   { $prop = $curr.PSObject.Properties[$step] } catch { return $null }
        $curr = if ($prop) { $prop.Value } else { return $null }
    }
    return $curr
}

function Get-GraphToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$Scope = 'https://graph.microsoft.com/.default'
    )

    # -- Client-credentials flow using a client secret (no certificate) ------
    # After the customer grants admin consent to this multi-tenant app, we can
    # mint an app-only Graph token scoped to THEIR tenant simply by passing
    # their TenantId here. Nothing is installed on the target tenant.
    $tokenBody = @{
        grant_type    = 'client_credentials'
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = $Scope
    }

    $response = Invoke-RestMethod `
        -Uri    "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -Method POST `
        -Body   $tokenBody
    return $response.access_token
}

#endregion

#region -- Prerequisites ------------------------------------------------------

if (-not (Test-Path $OutputFolder)) { New-Item $OutputFolder -ItemType Directory | Out-Null }

# Verify Exchange Online module
# GRAPH-ONLY MODE: Exchange Online app-only PowerShell requires a certificate,
# which this client-secret build intentionally does not use. Exchange-specific
# checks are therefore skipped and reported as MANUAL in the results.
$SkipExchange = $true

# Verify ImportExcel (for XLSX output - optional, falls back to CSV)
$HasImportExcel = [bool](Get-Module ImportExcel -ListAvailable)

#endregion

#region -- Authentication -----------------------------------------------------

Write-Section "Authenticating"

Write-Host "  Getting Microsoft Graph token (client secret)..." -NoNewline
$GraphToken    = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
$GraphHeaders  = @{ Authorization = "Bearer $GraphToken"; 'Content-Type' = 'application/json' }
Write-Host " OK" -ForegroundColor Green

# Fetch organisation info up front so it is available whether or not Exchange
# Online is connected (the Graph-only build skips Exchange but still needs $OrgInfo).
$OrgInfo = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization" -Headers $GraphHeaders

if (-not $SkipExchange) {
    Write-Host "  Connecting Exchange Online (certificate)..." -NoNewline

    # Resolve the default domain name into a plain string BEFORE passing it
    # to Connect-ExchangeOnline.  Doing it inline with a pipeline causes PS
    # to pipe the Connect-ExchangeOnline output instead of the inner expression.
    $ExoOrg       = ($OrgInfo.value[0].verifiedDomains | Where-Object { $_.isDefault } | Select-Object -First 1).name

    Connect-ExchangeOnline `
        -CertificateThumbprint $null `
        -AppId                 $ClientId `
        -Organization          $ExoOrg `
        -ShowBanner:$false `
        -ErrorAction           Stop

    Write-Host " OK ($ExoOrg)" -ForegroundColor Green
}

#endregion

#region -- Raw Data Collection ------------------------------------------------

Write-Section "Collecting raw data from Graph API"

# -- Organisation -------------------------------------------------------------
Write-Host "  [1/30] Organisation..." -NoNewline
# Reuse the response already fetched during authentication (saves one Graph call)
$Org           = $OrgInfo.value[0]
$DefaultDomain = ($Org.verifiedDomains | Where-Object { $_.isDefault } | Select-Object -First 1).name
Write-Host " $DefaultDomain" -ForegroundColor Green

# -- Secure Score -------------------------------------------------------------
Write-Host "  [2/30] Secure Score..." -NoNewline
try {
    $SecureScoreAll = (Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/security/secureScores?`$top=1" `
        -Headers $GraphHeaders).value
    $SecureScore    = $SecureScoreAll[0]
    $SecureScorePct = [math]::Round(($SecureScore.currentScore / $SecureScore.maxScore) * 100, 0)
    Write-Host " $($SecureScore.currentScore)/$($SecureScore.maxScore) ($SecureScorePct%)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED (SecurityEvents.Read.All missing - grant admin consent)" -ForegroundColor Yellow
    $SecureScore    = [PSCustomObject]@{ currentScore = 0; maxScore = 0 }
    $SecureScorePct = 0
}

# Secure Score history (30 days for trend chart)
try {
    $SecureScoreHistory = (Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/security/secureScores?`$top=30" `
        -Headers $GraphHeaders).value
} catch {
    $SecureScoreHistory = @()
}

# -- Users ---------------------------------------------------------------------
Write-Host "  [3/30] Users (all)..." -NoNewline
try {
    $AllUsers = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,accountEnabled,userType,signInActivity,passwordPolicies,createdDateTime&`$top=999" `
        -Headers $GraphHeaders
    Write-Host " $($AllUsers.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
    $AllUsers = @()
}

# Filter sets
$EnabledUsers = @($AllUsers | Where-Object { $_.accountEnabled -eq $true -and $_.userType -ne 'Guest' })
$GuestUsers   = @($AllUsers | Where-Object { $_.userType -eq 'Guest' })
$CutoffDate   = (Get-Date).AddDays(-$DormantDays)
$DormantUsers = @($EnabledUsers | Where-Object {
    $sia = $_.PSObject.Properties['signInActivity']
    # No signInActivity property at all, OR property is null, OR last sign-in is before cutoff
    (-not $sia) -or
    ($null -eq $sia.Value) -or
    (-not $sia.Value.lastSignInDateTime) -or
    ([datetime]$sia.Value.lastSignInDateTime -lt $CutoffDate)
})

# -- MFA Registration ---------------------------------------------------------
Write-Host "  [4/30] MFA registration report..." -NoNewline
try {
    $MfaReg = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?`$top=999" `
        -Headers $GraphHeaders
    Write-Host " $($MfaReg.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED (Reports.Read.All missing)" -ForegroundColor Yellow
    $MfaReg = @()
}

$NoMfaUsers   = @($MfaReg | Where-Object { -not $_.isMfaRegistered })
$WeakMfaUsers = @($MfaReg | Where-Object {
    $_.isMfaRegistered -and
    ($_.methodsRegistered -contains 'sms' -or
     $_.methodsRegistered -contains 'voice' -or
     $_.methodsRegistered -contains 'email') -and
    -not ($_.methodsRegistered -contains 'microsoftAuthenticatorPush' -or
          $_.methodsRegistered -contains 'fido2' -or
          $_.methodsRegistered -contains 'windowsHelloForBusiness' -or
          $_.methodsRegistered -contains 'softwareOneTimePasscode')
})

# -- Risky Users ---------------------------------------------------------------
# Requires IdentityRiskyUser.Read.All + Entra ID P2 (included in E5).
# Max page size for this endpoint is 100; filter locally after fetching.
Write-Host "  [5/30] Risky users..." -NoNewline
try {
    $AllRiskyUsers = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/identityProtection/riskyUsers?`$top=100" `
        -Headers $GraphHeaders
    $RiskyUsers = @($AllRiskyUsers | Where-Object {
        $rs = $_.PSObject.Properties['riskState']
        $rs -and $rs.Value -notin @('dismissed','remediated','none')
    })
    Write-Host " $($RiskyUsers.Count) active (of $($AllRiskyUsers.Count) total)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED - $($_.Exception.Message)" -ForegroundColor Yellow
    $RiskyUsers = @()
}

# -- Directory Roles -----------------------------------------------------------
Write-Host "  [6/30] Directory roles..." -NoNewline
try {
    $DirectoryRoles  = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/directoryRoles?`$expand=members" `
        -Headers $GraphHeaders
    $GlobalAdminRole = $DirectoryRoles | Where-Object { $_.displayName -eq 'Global Administrator' }
    $GlobalAdmins    = if ($GlobalAdminRole) { $GlobalAdminRole.members } else { @() }
    Write-Host " GlobalAdmins=$($GlobalAdmins.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
    $DirectoryRoles = @(); $GlobalAdmins = @()
}

# -- Conditional Access Policies -----------------------------------------------
Write-Host "  [7/30] Conditional Access Policies..." -NoNewline
try {
    $CAPolicies = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" `
        -Headers $GraphHeaders
    Write-Host " $($CAPolicies.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED (Policy.Read.All missing)" -ForegroundColor Yellow
    $CAPolicies = @()
}

# -- OAuth delegated consent grants (for 1.25) ----------------------------------
# Tenant-wide in one call - not the *policy* on future consent (that's 1.7),
# this is the actual inventory of what's already been granted. The classic
# illicit-consent-grant attack (malicious "Enable editing"-style OAuth apps)
# lives here, not in the consent policy setting.
Write-Host "  [+] OAuth delegated consent grants..." -NoNewline
$OAuth2Grants = @()
try {
    $OAuth2Grants = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" `
        -Headers $GraphHeaders
    Write-Host " $($OAuth2Grants.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
}

# -- Application permission grants on Microsoft Graph (for 1.25) ---------------
# One call against Graph's own service principal rather than iterating every
# app in the tenant (hundreds of calls) - covers the resource this attack
# pattern most commonly targets.
Write-Host "  [+] Application permission grants (Microsoft Graph)..." -NoNewline
$AppRoleGrants = @()
$GraphAppRoleNames = @{}
try {
    $AppRoleGrants = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='00000003-0000-0000-c000-000000000000')/appRoleAssignedTo" `
        -Headers $GraphHeaders
    Write-Host " $($AppRoleGrants.Count)" -ForegroundColor Green
    # Resolve appRoleId -> readable permission name (e.g. "Mail.ReadWrite")
    # for the grants just collected - one extra call, only worth making if
    # there's anything to resolve.
    if ($AppRoleGrants.Count -gt 0) {
        $graphSp = Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='00000003-0000-0000-c000-000000000000')?`$select=appRoles" `
            -Headers $GraphHeaders
        foreach ($role in @($graphSp.appRoles)) {
            $GraphAppRoleNames[[string]$role.id] = [string]$role.value
        }
    }
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
}

# -- Application registration credentials (for 1.26) ---------------------------
Write-Host "  [+] Application registration credentials..." -NoNewline
$AppRegistrations = @()
try {
    $AppRegistrations = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/applications?`$select=id,displayName,keyCredentials,passwordCredentials" `
        -Headers $GraphHeaders
    Write-Host " $($AppRegistrations.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
}

# -- Authorization Policy ------------------------------------------------------
Write-Host "  [8/30] Authorization policy..." -NoNewline
try {
    $AuthPolicy = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/policies/authorizationPolicy" `
        -Headers $GraphHeaders
    Write-Host " OK" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
    $AuthPolicy = [PSCustomObject]@{
        defaultUserRolePermissions = [PSCustomObject]@{ allowedToCreateApps = $null }
        guestUserRoleId            = $null
        permissionGrantPoliciesAssigned = @()
    }
}

# -- Password Expiry Policy ----------------------------------------------------
Write-Host "  [9/30] Password expiry domains..." -NoNewline
try {
    $Domains = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/domains?`$select=id,isVerified,isDefault,authenticationType,passwordNotificationWindowInDays,passwordValidityPeriodInDays" `
        -Headers $GraphHeaders
    Write-Host " $($Domains.Count) domains" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
    $Domains = @()
}

# -- Devices -------------------------------------------------------------------
Write-Host "  [10/30] Devices (Entra)..." -NoNewline
try {
    $EntraDevices = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/devices?`$select=id,displayName,operatingSystem,operatingSystemVersion,isCompliant,isManaged,approximateLastSignInDateTime,trustType&`$top=999" `
        -Headers $GraphHeaders
    Write-Host " $($EntraDevices.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
    $EntraDevices = @()
}

$StaleDate    = (Get-Date).AddDays(-$StaleDeviceDays)
$StaleDevices = @($EntraDevices | Where-Object {
    $lsi = $_.PSObject.Properties['approximateLastSignInDateTime']
    $lsi -and $lsi.Value -and [datetime]$lsi.Value -lt $StaleDate
})

# -- Intune - Device Compliance ------------------------------------------------
Write-Host "  [11/30] Intune managed devices..." -NoNewline
try {
    $IntuneDevices = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=id,deviceName,operatingSystem,osVersion,complianceState,isEncrypted,lastSyncDateTime,managedDeviceOwnerType&`$top=999" `
        -Headers $GraphHeaders
    Write-Host " $($IntuneDevices.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED (DeviceManagementManagedDevices.Read.All missing)" -ForegroundColor Yellow
    $IntuneDevices = @()
}

$NonCompliantDevices = @($IntuneDevices | Where-Object { $_.complianceState -eq 'noncompliant' })
$NotEncryptedDevices = @($IntuneDevices | Where-Object { $_.isEncrypted -eq $false })
$PersonalDevices     = @($IntuneDevices | Where-Object { $_.managedDeviceOwnerType -eq 'personal' })

# -- Intune - Compliance Policies ---------------------------------------------
Write-Host "  [12/30] Intune compliance policies..." -NoNewline
try {
    $CompliancePolicies = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies" `
        -Headers $GraphHeaders
    Write-Host " $($CompliancePolicies.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED (DeviceManagementConfiguration.Read.All missing)" -ForegroundColor Yellow
    $CompliancePolicies = @()
}

# -- Intune - Configuration Profiles ------------------------------------------
Write-Host "  [13/30] Intune configuration profiles..." -NoNewline
try {
    $ConfigProfiles = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations" `
        -Headers $GraphHeaders
    Write-Host " $($ConfigProfiles.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
    $ConfigProfiles = @()
}

# -- Intune - Software Update Rings --------------------------------------------
Write-Host "  [14/30] Intune update rings..." -NoNewline
try {
    $UpdateRings = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations?`$filter=isof('microsoft.graph.windowsUpdateForBusinessConfiguration')" `
        -Headers $GraphHeaders
    Write-Host " $($UpdateRings.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
    $UpdateRings = @()
}

# -- Intune - App Protection Policies -----------------------------------------
Write-Host "  [15/30] App protection policies..." -NoNewline
try {
    $AppProtectioniOS     = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/iosManagedAppProtections"    -Headers $GraphHeaders).value
    $AppProtectionAndroid = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/androidManagedAppProtections" -Headers $GraphHeaders).value
    Write-Host " iOS=$($AppProtectioniOS.Count) Android=$($AppProtectionAndroid.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
    $AppProtectioniOS = @(); $AppProtectionAndroid = @()
}

# -- Intune - Security Baselines -----------------------------------------------
# /deviceManagement/intents is beta-only; v1.0 returns 400.
# Use the beta endpoint and filter locally for assigned baselines.
Write-Host "  [16/30] Security baseline profiles..." -NoNewline
try {
    $AllIntents        = @(Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/beta/deviceManagement/intents" `
        -Headers $GraphHeaders)
    $SecurityBaselines = @($AllIntents | Where-Object {
        $a = $_.PSObject.Properties['isAssigned']
        $a -and $a.Value -eq $true
    })
    Write-Host " $($SecurityBaselines.Count) assigned (of $($AllIntents.Count) total)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
    $SecurityBaselines = @()
}

# -- Intune - LAPS -------------------------------------------------------------
Write-Host "  [17/30] LAPS configuration..." -NoNewline
# Wrap in @() - strict mode throws .Count on a single object returned by Where-Object
$LapsPolicy = @($ConfigProfiles | Where-Object {
    $_.displayName -match 'LAPS|Local Admin' -or
    ($_.PSObject.Properties['@odata.type'] -and
     $_.PSObject.Properties['@odata.type'].Value -match 'laps')
})
Write-Host " $($LapsPolicy.Count) matching profiles" -ForegroundColor Green

# -- Enterprise Applications ----------------------------------------------------
Write-Host "  [18/30] Enterprise applications (service principals)..." -NoNewline
try {
    $EnterpriseApps = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$select=id,displayName,appId,publisherName,signInAudience,tags&`$top=999" `
        -Headers $GraphHeaders
    Write-Host " $($EnterpriseApps.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
    $EnterpriseApps = @()
}

# -- Dynamic Groups -------------------------------------------------------------
Write-Host "  [19/30] Dynamic groups..." -NoNewline
try {
    $DynamicGroups = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(c:c eq 'DynamicMembership')&`$select=id,displayName" `
        -Headers $GraphHeaders
    Write-Host " $($DynamicGroups.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
    $DynamicGroups = @()
}

# -- SharePoint Tenant Settings ------------------------------------------------
# Requires: SharePointTenantSettings.Read.All (application permission).
# If absent, the script reports 'unknown' for sharing settings - the assessment
# and report will still complete successfully, but check 5.1 cannot be automated.
Write-Host "  [20/30] SharePoint tenant settings..." -NoNewline
try {
    $SPSettings = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/admin/sharepoint/settings" `
        -Headers $GraphHeaders
    Write-Host " SharingCapability=$($SPSettings.sharingCapability)" -ForegroundColor Green
} catch {
    try {
        # Beta fallback - same endpoint, sometimes more permissive
        $SPSettings = Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/beta/admin/sharepoint/settings" `
            -Headers $GraphHeaders
        Write-Host " SharingCapability=$($SPSettings.sharingCapability) (beta)" -ForegroundColor Green
    } catch {
        $errMsg = $_.Exception.Message
        if ($errMsg -match '403|Forbidden') {
            Write-Host " SKIPPED (403 Forbidden - grant SharePointTenantSettings.Read.All to retrieve)" -ForegroundColor Yellow
        } else {
            Write-Host " SKIPPED ($errMsg)" -ForegroundColor Yellow
        }
        $SPSettings = [PSCustomObject]@{
            sharingCapability      = 'unknown'
            defaultSharingLinkType = 'unknown'
            defaultLinkPermission  = 'unknown'
        }
    }
}

# -- SharePoint Sites ---------------------------------------------------------
# NOTE: `/sites?search=*` does NOT return every site in the tenant - it only
# returns sites the search index can match and silently omits many (this is why
# earlier runs "couldn't retrieve all SharePoint sites"). The dedicated
# enumeration endpoint `/sites/getAllSites` (beta) returns the COMPLETE list.
# We try it first, then fall back to search, then to the root-site delta.
Write-Host "  [21/30] SharePoint sites..." -NoNewline
$SPSites = @()
$spMethod = ''
try {
    # Full-tenant enumeration (proper API). Requires Sites.Read.All.
    $SPSites = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/beta/sites/getAllSites?`$select=id,displayName,name,webUrl,isPersonalSite" `
        -Headers $GraphHeaders
    $spMethod = 'getAllSites'
} catch {
    try {
        # Fallback 1: search index (partial, but broad)
        $SPSites = Invoke-GraphAll `
            -Uri "https://graph.microsoft.com/v1.0/sites?search=*&`$top=200&`$select=id,displayName,name,webUrl" `
            -Headers $GraphHeaders
        $spMethod = 'search=*'
    } catch {
        try {
            # Fallback 2: enumerate all sites without a search term (v1.0)
            $SPSites = Invoke-GraphAll `
                -Uri "https://graph.microsoft.com/v1.0/sites?`$top=200&`$select=id,displayName,name,webUrl" `
                -Headers $GraphHeaders
            $spMethod = 'sites-list'
        } catch {
            Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
            $SPSites = @()
        }
    }
}
# Exclude personal OneDrive sites from the site-collection counts where flagged
if ($SPSites.Count -gt 0) {
    $SPSites = @($SPSites | Where-Object {
        $ips = $_.PSObject.Properties['isPersonalSite']
        -not ($ips -and $ips.Value -eq $true)
    })
    Write-Host " $($SPSites.Count) (via $spMethod)" -ForegroundColor Green
}

# -- Defender Secure Score Controls --------------------------------------------
Write-Host "  [22/30] Secure score control profiles..." -NoNewline
try {
    $SecureScoreControls = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/security/secureScoreControlProfiles?`$top=999" `
        -Headers $GraphHeaders
    Write-Host " $($SecureScoreControls.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED (SecurityEvents.Read.All missing)" -ForegroundColor Yellow
    $SecureScoreControls = @()
}

$ControlMap = @{}
foreach ($c in $SecureScoreControls) {
    $cn = $c.PSObject.Properties['controlName']
    if ($cn -and $cn.Value) { $ControlMap[$cn.Value] = $c }
}

# -- PIM (Privileged Identity Management) -------------------------------------
Write-Host "  [23/30] PIM role assignments (eligible)..." -NoNewline
try {
    $PimEligible = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?`$top=999" `
        -Headers $GraphHeaders
    Write-Host " $($PimEligible.Count)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED (PIM not licensed or RoleManagement.Read.Directory missing)" -ForegroundColor Yellow
    $PimEligible = @()
}

# -- Defender MDE Devices ------------------------------------------------------
Write-Host "  [24/30] Defender for Endpoint machines..." -NoNewline
$MdeDevices       = @()   # Requires Defender API (separate token scope); skipped here
$DefenderEnrolled = @($IntuneDevices | Where-Object { $_.managedDeviceOwnerType -ne $null })
Write-Host " Using Intune proxy ($($DefenderEnrolled.Count) Windows devices)" -ForegroundColor Green


# -- Licensing / Subscribed SKUs ----------------------------------------------
Write-Host "  [+] Licensing (subscribed SKUs)..." -NoNewline
$SubscribedSkus = @()
try {
    # subscribedSkus does not support $top - passing it causes an
    # unconditional 400 Bad Request regardless of tenant size or
    # permissions. Confirmed live: with $top=999 this always failed;
    # without it, the call succeeds (this endpoint is never large enough
    # to need paging - it's one row per purchased licence SKU). This means
    # the Licensing Overview report page has been broken/empty on every
    # real customer report generated by this pipeline until now.
    $SubscribedSkus = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" `
        -Headers $GraphHeaders
    Write-Host " $($SubscribedSkus.Count) SKU(s)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
}

# -- Mail Activity Report (Graph Reports API) --------------------------------
Write-Host "  [+] Email activity report (last 30 days)..." -NoNewline
$MailActivityReport = @()
try {
    $maUri  = "https://graph.microsoft.com/v1.0/reports/getEmailActivityCounts(period='D30')"
    $maCsv  = Invoke-RestMethod -Uri $maUri -Headers $GraphHeaders -ErrorAction Stop
    # Graph returns CSV text; parse it
    $maLines = ($maCsv -split "`n" | Where-Object { $_.Trim() -ne '' })
    if ($maLines.Count -gt 1) {
        $MailActivityReport = $maLines | Select-Object -Skip 1 | ConvertFrom-Csv -Header ($maLines[0] -split ',')
    }
    Write-Host " $($MailActivityReport.Count) day(s)" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
}

# -- Authentication Methods Policy --------------------------------------------
Write-Host "  [+] Authentication methods policy..." -NoNewline
$AuthMethodsPolicy = $null
try {
    $AuthMethodsPolicy = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy" `
        -Headers $GraphHeaders -ErrorAction Stop
    Write-Host " OK" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
}

# -- Cross-Tenant Access / External Collaboration Settings --------------------
Write-Host "  [+] External collaboration settings..." -NoNewline
$ExternalCollabSettings = $null
try {
    $ExternalCollabSettings = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy" `
        -Headers $GraphHeaders -ErrorAction Stop
    Write-Host " OK" -ForegroundColor Green
} catch {
    Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
}

# -- Mailboxes (Exchange) ------------------------------------------------------
if (-not $SkipExchange) {
    Write-Host "  [25/30] Mailboxes (Exchange Online)..." -NoNewline
    $Mailboxes = Get-Mailbox -ResultSize Unlimited -Filter { RecipientTypeDetails -ne 'SharedMailbox' }
    Write-Host " $($Mailboxes.Count)" -ForegroundColor Green

    # -- Transport Rules ------------------------------------------------------
    Write-Host "  [26/30] Transport rules..." -NoNewline
    $TransportRules = Get-TransportRule
    Write-Host " $($TransportRules.Count)" -ForegroundColor Green

    # -- Inbox Rules ---------------------------------------------------------
    # Sampling only - one Get-InboxRule call per mailbox takes 2-4 s each.
    # Default sample of 50 keeps this step under ~3 minutes.
    # Raise $InboxRuleSample at the top of the script if you want broader coverage.
    $InboxRuleSample = 5
    Write-Host "  [27/30] Inbox rules (first $InboxRuleSample mailboxes)..." -NoNewline
    $InboxRules = @()
    $mbSample   = $Mailboxes | Select-Object -First $InboxRuleSample
    $mbDone     = 0
    foreach ($mb in $mbSample) {
        try {
            $rules = @(Get-InboxRule -Mailbox $mb.PrimarySmtpAddress `
                        -ErrorAction SilentlyContinue `
                        -WarningAction SilentlyContinue)
            $InboxRules += $rules
        } catch {}
        $mbDone++
        if ($mbDone % 10 -eq 0) {
            Write-Host "." -NoNewline   # progress dot every 10 mailboxes
        }
    }
    Write-Host " $($InboxRules.Count) rules (sampled $InboxRuleSample mailboxes)" -ForegroundColor Green

    # -- Anti-Phishing --------------------------------------------------------
    Write-Host "  [28/30] Anti-phishing policies..." -NoNewline
    $AntiPhishPolicies = Get-AntiPhishPolicy
    Write-Host " $($AntiPhishPolicies.Count)" -ForegroundColor Green

    # -- Safe Links / Safe Attachments ----------------------------------------
    Write-Host "  [29/30] Safe Links / Safe Attachments..." -NoNewline
    $SafeLinksPolicies      = Get-SafeLinksPolicy
    $SafeAttachmentPolicies = Get-SafeAttachmentPolicy
    Write-Host " SafeLinks=$($SafeLinksPolicies.Count) SafeAtt=$($SafeAttachmentPolicies.Count)" -ForegroundColor Green

    # -- DKIM signing configuration (for Email Health report page) -------------
    Write-Host "  [+] DKIM signing configuration..." -NoNewline
    try {
        $DkimConfigs = @(Get-DkimSigningConfig -ErrorAction Stop)
        Write-Host " $($DkimConfigs.Count) domains" -ForegroundColor Green
    } catch {
        Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
        $DkimConfigs = @()
    }

    # -- Mail flow statistics - last 30 days (for Email Health report page) -----
    Write-Host "  [+] Mail flow statistics (30 days)..." -NoNewline
    try {
        $mfStart  = (Get-Date).AddDays(-30)
        $mfEnd    = Get-Date
        $mfReport = @(Get-MailflowStatusReport -StartDate $mfStart -EndDate $mfEnd -ErrorAction Stop)
        Write-Host " $($mfReport.Count) rows" -ForegroundColor Green
    } catch {
        Write-Host " SKIPPED ($($_.Exception.Message))" -ForegroundColor Yellow
        $mfReport = @()
    }

}

#endregion

#region -- SECTION 1: Entra ID -----------------------------------------------
Write-Section "1 - Entra ID"

# Helper: find CA policies matching a predicate
function Find-CaPolicy {
    param([ScriptBlock]$Filter)
    $CAPolicies | Where-Object { $_.state -eq 'enabled' } | Where-Object $Filter
}

# - 1.1 MFA enforced for ALL users --------------------------------------------
$mfaAllPolicy = Find-CaPolicy {
    (Get-Prop $_ 'grantControls.builtInControls') -contains 'mfa' -and
    (Get-Prop $_ 'conditions.users.includeUsers')  -contains 'All'
}
$status_1_1a = if ($mfaAllPolicy) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.1' 'MFA is enforced for all users' $status_1_1a 'critical' `
    $(if (-not $mfaAllPolicy) {'A conditional access policy is either missing or misconfigured.'} else {'Conditional Access Policy enforces MFA for all users.'})

# Azure Management MFA
$mfaAzureMgmt = Find-CaPolicy {
    (Get-Prop $_ 'conditions.applications.includeApplications') -contains 'MicrosoftAzureManagement' -and
    (Get-Prop $_ 'grantControls.builtInControls') -contains 'mfa'
}
$status_1_1b = if ($mfaAzureMgmt) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.1' 'MFA is enforced for Azure Management' $status_1_1b 'critical' `
    $(if (-not $mfaAzureMgmt) {'A conditional access policy is either missing or misconfigured.'} else {'Policy found.'})

# MFA Enrollment coverage
$noMfaCount   = $NoMfaUsers.Count
$status_1_1c  = if ($noMfaCount -eq 0) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.1' 'Users are enrolled in MFA and covered by a policy' $status_1_1c 'critical' `
    $(if ($noMfaCount -gt 0) {"$noMfaCount users do not have MFA enabled"} else {'All users have MFA registered.'})

Write-Check "1.1 MFA for all users"       $status_1_1a
Write-Check "1.1 MFA for Azure Mgmt"      $status_1_1b
Write-Check "1.1 MFA enrollment coverage" $status_1_1c "$noMfaCount users without MFA"

# - 1.2 MFA required for Admins ------------------------------------------------
$mfaAdminPolicy = Find-CaPolicy {
    @(Get-Prop $_ 'conditions.users.includeRoles').Count -gt 0 -and
    (Get-Prop $_ 'grantControls.builtInControls') -contains 'mfa'
}
$status_1_2 = if ($mfaAdminPolicy) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.2' 'MFA is enforced on accounts with highly privileged roles' $status_1_2 'critical' `
    $(if ($mfaAdminPolicy) {'Conditional Access Policy found that is enforcing MFA for admins.'} else {'No admin MFA CA policy found.'})
Write-Check "1.2 MFA required for Admins" $status_1_2

# - 1.3 Legacy Authentication blocked ------------------------------------------
$legacyBlock = Find-CaPolicy {
    $c = Get-Prop $_ 'conditions.clientAppTypes'
    ($c -contains 'exchangeActiveSync' -or $c -contains 'other') -and
    (Get-Prop $_ 'grantControls.builtInControls') -contains 'block'
}
$status_1_3 = if ($legacyBlock) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.3' 'Legacy Authentication shall be blocked' $status_1_3 'high' `
    $(if ($legacyBlock) {'Legacy auth block policy found.'} else {'No conditional access policy found'})
Write-Check "1.3 Legacy Authentication blocked" $status_1_3

# - 1.4 Break Glass accounts (MANUAL) ------------------------------------------
# There's no reliable way to identify which accounts ARE break-glass from
# Graph data alone (no directory flag for this), so this stays MANUAL - but
# now surfaces the actual CA exclusion list instead of a bare "Manual."
# placeholder, giving the reviewer something concrete to check against.
$excludedUserIds = @($CAPolicies | ForEach-Object { @(Get-Prop $_ 'conditions.users.excludeUsers') } |
    Where-Object { $_ -and $_ -ne 'GuestsOrExternalUsers' } | Select-Object -Unique)
$breakGlassDetail = if ($excludedUserIds.Count -gt 0) {
    "$($excludedUserIds.Count) user(s) are excluded from one or more Conditional Access policies - confirm these are your break-glass/emergency-access accounts, and that they're excluded from every CA policy, not just some."
} else {
    "No users are excluded from any Conditional Access policy. If you don't have dedicated break-glass accounts intentionally excluded, a CA misconfiguration or outage could lock out every administrator at once."
}
Add-Result '1-EntraID' '1.4' 'Break Glass users are created for emergency access' 'MANUAL' 'high' $breakGlassDetail
Write-Check "1.4 Break Glass accounts" 'MANUAL'

# - 1.5 Global Admins 2-4 ------------------------------------------------------
$gaCount    = $GlobalAdmins.Count
$status_1_5 = if ($gaCount -ge 2 -and $gaCount -le 4) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.5' 'Ensure that between two and four global admins are designated' $status_1_5 'high' `
    "$gaCount Global admin$(if ($gaCount -ne 1) {'s'}) were detected."
Write-Check "1.5 Global Admin count (2-4)" $status_1_5 "Current: $gaCount"

# - 1.6 Global Admins cloud-only -----------------------------------------------
$syncedAdmins  = @($GlobalAdmins | Where-Object { $_.onPremisesSyncEnabled -eq $true })
$status_1_6    = if ($syncedAdmins.Count -eq 0) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.6' 'Ensure Administrative accounts are cloud-only' $status_1_6 'high' `
    $(if ($status_1_6 -eq 'PASS') {'All Global Admins are cloud-only.'} else {"$($syncedAdmins.Count) synced Global Admin(s) found."})
Write-Check "1.6 Global Admins cloud-only" $status_1_6

# - 1.7 User consent / 3rd party app registration -----------------------------
$userConsentOk = (Get-Prop $AuthPolicy 'defaultUserRolePermissions.allowedToCreateApps') -eq $false
$status_1_7a   = if ($userConsentOk) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.7' 'Only Admins shall be allowed to register 3rd party applications' $status_1_7a 'high' `
    $(if ($userConsentOk) {'Authorization Policy configured correctly.'} else {'Authorization Policy - users can register apps.'})

$pgpa        = @(Get-Prop $AuthPolicy 'permissionGrantPoliciesAssigned')
$status_1_7b = if ($pgpa.Count -eq 0 -or $pgpa -notcontains 'ManagePermissionGrantsForSelf.microsoft-user-default-legacy-v2') {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.7' 'Non-admin users shall be prevented from providing consent to 3rd party applications' $status_1_7b 'medium' `
    $(if ($status_1_7b -eq 'PASS') {'Authorization Policy configured correctly.'} else {'Authorization Policy - user consent not blocked.'})

Write-Check "1.7 App registration (admins only)" $status_1_7a
Write-Check "1.7 User consent blocked"            $status_1_7b

# - 1.8 Guest user restricted access ------------------------------------------
$guestRestricted = $AuthPolicy.guestUserRoleId -in @(
    '2af84b1e-32c8-42b7-82bc-daa82404023b',  # Restricted Guest
    '10dae51f-b6af-4016-8d66-8c2a99b929b3'   # Guest
)
$status_1_8 = if ($guestRestricted) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.8' 'Guest users have limited access to properties and memberships of directory objects' $status_1_8 'low' `
    $(if ($status_1_8 -eq 'PASS') {'Guest users have limited access to directory objects.'} else {'Guest access is not restricted.'})
Write-Check "1.8 Guest restricted access" $status_1_8

# - 1.9 Passwords do not expire ------------------------------------------------
$expiringDomains = @($Domains | Where-Object {
    $pvp = $_.PSObject.Properties['passwordValidityPeriodInDays']
    $pvp -and $pvp.Value -and $pvp.Value -ne 2147483647
})
$status_1_9 = if ($expiringDomains.Count -eq 0) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.9' 'Passwords shall not expire' $status_1_9 'medium' `
    $(if ($status_1_9 -eq 'PASS') {'Passwords are set to never expire.'} else {'Passwords expire.'})
Write-Check "1.9 Passwords do not expire" $status_1_9

# - 1.10 MFA required to enroll devices ----------------------------------------
$mfaDeviceEnroll = Find-CaPolicy {
    (Get-Prop $_ 'conditions.applications.includeUserActions') -contains 'urn:user:registersecurityinfo' -and
    (Get-Prop $_ 'grantControls.builtInControls') -contains 'mfa'
}
$status_1_10 = if ($mfaDeviceEnroll) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.10' 'MFA shall be required to enroll devices to Entra ID' $status_1_10 'low' `
    $(if ($mfaDeviceEnroll) {'Policy found.'} else {'A conditional access policy is either missing or misconfigured.'})
Write-Check "1.10 MFA for device enrollment" $status_1_10

# - 1.11 Local Administrator settings for device joins ------------------------
$deviceJoinCa = Find-CaPolicy {
    (Get-Prop $_ 'conditions.applications.includeUserActions') -contains 'urn:user:registerdevice' -and
    (Get-Prop $_ 'grantControls.builtInControls') -contains 'mfa'
}
$localAdminSettings = $null
try {
    $localAdminSettings = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceRegistrationPolicy" `
        -Headers $GraphHeaders -ErrorAction Stop
} catch { <# endpoint unavailable or permissions missing - treat as not configured #> }
$status_1_11 = if ((Get-Prop $localAdminSettings 'localAdminPassword.isEnabled') -eq $true) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.11' 'Local Administrator settings are configured for device joins' $status_1_11 'medium' `
    $(if ($status_1_11 -eq 'PASS') {'Local admin password settings configured.'} else {'Local Administrator settings are not configured for device joins.'})
Write-Check "1.11 Local Admin settings for device joins" $status_1_11

# - 1.12 Dormant accounts disabled after 45 days ------------------------------
$dormantCount = $DormantUsers.Count
$status_1_12  = if ($dormantCount -eq 0) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.12' 'Dormant accounts are disabled after 45 days' $status_1_12 'medium' `
    "$dormantCount accounts were found active that have not signed in for over $DormantDays days."
Write-Check "1.12 Dormant accounts" $status_1_12 "$dormantCount dormant"

# - 1.13 Browser sessions not persistent for privileged users ------------------
$noPersBrowserPolicy = Find-CaPolicy {
    @(Get-Prop $_ 'conditions.users.includeRoles').Count -gt 0 -and
    (Get-Prop $_ 'sessionControls.persistentBrowser.isEnabled') -eq $true -and
    (Get-Prop $_ 'sessionControls.persistentBrowser.mode') -eq 'never'
}
$status_1_13 = if ($noPersBrowserPolicy) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.13' 'Browser Sessions shall not be persistent for privileged users' $status_1_13 'medium' `
    $(if ($noPersBrowserPolicy) {'Policy found.'} else {'No conditional access policy found.'})
Write-Check "1.13 Browser session persistence" $status_1_13

# - 1.14 Stale devices deleted (30 days) --------------------------------------
$staleCount   = $StaleDevices.Count
$status_1_14  = if ($staleCount -eq 0) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.14' 'Devices shall be deleted that haven''t checked in for over 45 days.' $status_1_14 'medium' `
    "$staleCount Devices have not checked in for $StaleDeviceDays+ days"
Write-Check "1.14 Stale devices" $status_1_14 "$staleCount stale"

# - 1.15 Enterprise apps catalogued -------------------------------------------
$appCount   = $EnterpriseApps.Count
$status_1_15 = if ($appCount -gt 0) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.15' 'All corporate approved applications are cataloged and periodically reviewed' $status_1_15 'low' `
    "$appCount Enterprise applications were detected."
Write-Check "1.15 Enterprise apps catalogued" $status_1_15 "$appCount apps"

# - 1.16 Dynamic groups --------------------------------------------------------
$dynCount   = $DynamicGroups.Count
$status_1_16 = if ($dynCount -gt 0) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.16' 'Dynamic Groups are leveraged for automated group management' $status_1_16 'low' `
    $(if ($dynCount -gt 0) {'Dynamic Group(s) detected.'} else {'No dynamic groups found.'})
Write-Check "1.16 Dynamic groups" $status_1_16 "$dynCount groups"

# - 1.17 MFA for Intune Enrollment --------------------------------------------
$mfaIntunePolicy = Find-CaPolicy {
    (Get-Prop $_ 'conditions.applications.includeApplications') -contains 'MicrosoftIntune' -and
    (Get-Prop $_ 'grantControls.builtInControls') -contains 'mfa'
}
$status_1_17 = if ($mfaIntunePolicy) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.17' 'MFA Shall be required for Intune Enrollment' $status_1_17 'low' `
    $(if ($mfaIntunePolicy) {'Policy found.'} else {'A conditional access policy is either missing or misconfigured.'})
Write-Check "1.17 MFA for Intune enrollment" $status_1_17

# - 1.18 Managed Devices required for sign-in ---------------------------------
$managedDevicePolicy = Find-CaPolicy {
    $gc = Get-Prop $_ 'grantControls.builtInControls'
    $gc -contains 'compliantDevice' -or $gc -contains 'domainJoinedDevice'
}
$status_1_18 = if ($managedDevicePolicy) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.18' 'Managed Devices shall be required for authentication' $status_1_18 'medium' `
    $(if ($managedDevicePolicy) {'Policy found.'} else {'No conditional access policy found.'})
Write-Check "1.18 Managed devices for sign-in" $status_1_18

# - 1.19 Device compliance required -------------------------------------------
$compliancePolicy = Find-CaPolicy {
    (Get-Prop $_ 'grantControls.builtInControls') -contains 'compliantDevice'
}
$status_1_19 = if ($compliancePolicy) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.19' 'Noncompliant devices shall be blocked from accessing corporate resources.' $status_1_19 'medium' `
    $(if ($compliancePolicy) {'Policy found.'} else {'No conditional access policy found or misconfigured'})
Write-Check "1.19 Device compliance required" $status_1_19

# - 1.20 Phishing-resistant MFA for Admins ------------------------------------
$phishMfaPolicy = Find-CaPolicy {
    @(Get-Prop $_ 'conditions.users.includeRoles').Count -gt 0 -and
    ((Get-Prop $_ 'grantControls.authenticationStrength.requirementsSatisfied') -match 'mfa' -or
     (Get-Prop $_ 'grantControls.authenticationStrength.allowedCombinations')   -match 'fido2')
}
$status_1_20 = if ($phishMfaPolicy) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.20' "Ensure 'Phishing-resistant MFA strength' is required for Administrators" $status_1_20 'low' `
    $(if ($phishMfaPolicy) {'Phishing-resistant MFA policy found.'} else {'MFA used for authenticating administrators is not phishing resistant'})
Write-Check "1.20 Phishing-resistant MFA for Admins" $status_1_20

# - 1.21 High/medium risk sign-ins blocked ------------------------------------
$riskBlockPolicy = Find-CaPolicy {
    $srl = Get-Prop $_ 'conditions.signInRiskLevels'
    ($srl -contains 'high' -or $srl -contains 'medium') -and
    (Get-Prop $_ 'grantControls.builtInControls') -contains 'block'
}
$status_1_21 = if ($riskBlockPolicy) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.21' "Ensure 'sign-in risk' is blocked for medium and high risk" $status_1_21 'high' `
    $(if ($riskBlockPolicy) {'Risk-based block policy found.'} else {'No conditional access policy found'})
Write-Check "1.21 Risk-based sign-in block" $status_1_21

# - 1.22 PIM configured --------------------------------------------------------
$pimConfigured = $PimEligible.Count -gt 0
$status_1_22a  = if ($pimConfigured) {'PASS'} else {'MANUAL'}
Add-Result '1-EntraID' '1.22' "Ensure 'Privileged Identity Management' is used to manage roles" $status_1_22a 'high' `
    $(if ($pimConfigured) {'PIM eligible assignments detected.'} else {'Not Set'})
Add-Result '1-EntraID' '1.22' 'Ensure approval is required for Global Administrator role activation' 'MANUAL' 'high' 'Manual.'
Write-Check "1.22 PIM configured"    $status_1_22a
Write-Check "1.22 PIM GA approval"   'MANUAL'

# - 1.23 Sentinel ingest -------------------------------------------------------
Add-Result '1-EntraID' '1.23' 'Microsoft Sentinel shall be configured to ingest log information' 'MANUAL' 'medium' 'Manual.'
Write-Check "1.23 Sentinel ingest" 'MANUAL'

# - 1.24 Device code sign-in flow blocked -------------------------------------
$deviceCodeBlock = Find-CaPolicy {
    (Get-Prop $_ 'conditions.authenticationFlows.transferMethods') -contains 'deviceCodeFlow' -and
    (Get-Prop $_ 'grantControls.builtInControls') -contains 'block'
}
$status_1_24 = if ($deviceCodeBlock) {'PASS'} else {'FAIL'}
Add-Result '1-EntraID' '1.24' 'Ensure the device code sign-in flow is blocked' $status_1_24 'medium' `
    $(if ($deviceCodeBlock) {'Device code flow block policy found.'} else {'No conditional access policy found.'})
Write-Check "1.24 Device code flow blocked" $status_1_24

# - 1.25 High-risk OAuth consent grants are reviewed ----------------------------
# Inventories what's ALREADY been consented to (delegated + application
# permission grants), not the policy governing future consent (that's 1.7) -
# the two are genuinely different: a tenant can have 1.7 locked down today
# while still carrying years of legacy high-risk grants from before that
# policy existed, or from admin consent. This is the classic illicit-
# consent-grant attack surface (e.g. malicious "Enable editing"-style OAuth
# phishing apps).
$highRiskScopeNames = @(
    'Mail.ReadWrite', 'Mail.Send', 'Mail.Read',
    'Files.ReadWrite.All', 'Files.Read.All',
    'Sites.ReadWrite.All', 'Sites.FullControl.All',
    'Directory.ReadWrite.All', 'Directory.AccessAsUser.All',
    'User.ReadWrite.All', 'Group.ReadWrite.All',
    'Application.ReadWrite.All', 'RoleManagement.ReadWrite.Directory',
    'full_access_as_app', 'Contacts.ReadWrite'
)
$riskyDelegated = @($OAuth2Grants | Where-Object {
    $grantScopes = ("$($_.scope)" -split '\s+')
    @($grantScopes | Where-Object { $highRiskScopeNames -contains $_ }).Count -gt 0
})
$riskyAppGrants = @($AppRoleGrants | Where-Object {
    $highRiskScopeNames -contains $GraphAppRoleNames[[string]$_.appRoleId]
})
$totalRisky1_25 = $riskyDelegated.Count + $riskyAppGrants.Count
$status_1_25 = if ($totalRisky1_25 -eq 0) { 'PASS' } else { 'FAIL' }
$detail_1_25 = if ($totalRisky1_25 -eq 0) {
    'No high-risk delegated or application permission grants found on 3rd-party applications.'
} else {
    "$totalRisky1_25 high-risk grant(s) found ($($riskyDelegated.Count) delegated, $($riskyAppGrants.Count) application-level) among your tenant's consented apps - review Enterprise Applications for API permissions that go beyond what each app genuinely needs."
}
Add-Result '1-EntraID' '1.25' 'High-risk OAuth application permission grants are reviewed' $status_1_25 'high' $detail_1_25
Write-Check "1.25 High-risk OAuth grants" $status_1_25

# - 1.26 Application registration credential hygiene ----------------------------
$longLivedThresholdDays = 365   # CIS-style guidance: secrets/certs should be
                                 # rotated at least annually
$riskyCreds = New-Object System.Collections.Generic.List[object]
foreach ($app in $AppRegistrations) {
    foreach ($cred in (@(Get-Prop $app 'passwordCredentials') + @(Get-Prop $app 'keyCredentials'))) {
        if (-not $cred) { continue }
        $credStart = Get-Prop $cred 'startDateTime'
        $credEnd   = Get-Prop $cred 'endDateTime'
        if ($credStart -and $credEnd) {
            $lifetimeDays = ([datetime]$credEnd - [datetime]$credStart).TotalDays
            if ($lifetimeDays -gt $longLivedThresholdDays) {
                [void]$riskyCreds.Add([PSCustomObject]@{ App = [string]$app.displayName; LifetimeDays = [int]$lifetimeDays })
            }
        }
    }
}
$status_1_26 = if ($riskyCreds.Count -eq 0) { 'PASS' } else { 'FAIL' }
$riskyAppNames1_26 = @($riskyCreds | ForEach-Object { $_.App } | Select-Object -Unique)
$detail_1_26 = if ($riskyCreds.Count -eq 0) {
    'No application registration secrets or certificates exceed a 12-month lifetime.'
} else {
    "$($riskyCreds.Count) credential(s) across $($riskyAppNames1_26.Count) app registration(s) have a lifetime over 12 months - long-lived secrets are a standing risk if ever leaked. Affected apps: $($riskyAppNames1_26 -join ', ')."
}
Add-Result '1-EntraID' '1.26' 'Application registration credentials are not excessively long-lived' $status_1_26 'medium' $detail_1_26
Write-Check "1.26 App credential hygiene" $status_1_26

#endregion

#region -- SECTION 2: Exchange Online ----------------------------------------
Write-Section "2 - Exchange Online"

# - 2.1 SPF records -------------------------------------------------------
# Deliberately OUTSIDE the SkipExchange gate below: this is a plain public
# DNS lookup (Resolve-DnsName), not an Exchange Online PowerShell cmdlet -
# it needs no certificate, no Exchange.ManageAsApp, no extra consent at all,
# so there's no real reason to withhold it from this (Essential/free) tier.
# Confirmed: 2.3-2.6 below genuinely need Get-AntiPhishPolicy/
# Get-ExternalInOutlook/Get-TransportRule/Get-OrganizationConfig and stay
# tier-gated; 2.1/2.2 don't.
#
# Excludes *.onmicrosoft.com - Microsoft owns that DNS zone, so neither the
# customer nor Camelot can ever publish a record there. Without this,
# 2.2 (DMARC) was structurally impossible to pass for ANY tenant: Microsoft
# never publishes a DMARC record for the default onmicrosoft.com domain, so
# it always counted as "missing" regardless of what the customer's real
# domains had configured. Confirmed live against Camelot's own tenant
# 2026-09-14 - cdgiq.com's DMARC record was live and correct, yet 2.2 still
# failed on the onmicrosoft.com domain alone. Matches the same isOnMs
# exclusion the Email Health table below already applies (reported as 'NA'
# there, not PASS/FAIL) - this was simply never carried over to these two
# tenant-wide pass/fail checks.
$CheckableDomains = @($Domains | Where-Object { $_.id -notmatch '\.onmicrosoft\.com$' })

$domainsNoSpf = @($CheckableDomains | Where-Object {
    $d = $_
    try { -not (Resolve-DnsName -Name $d.id -Type TXT -ErrorAction Stop |
          Where-Object { $_.Text -match 'v=spf1' }) } catch { $true }
})
$status_2_1 = if ($domainsNoSpf.Count -eq 0) {'PASS'} else {'FAIL'}
Add-Result '2-Exchange' '2.1' 'SPF records shall be configured for all domains' $status_2_1 'high' `
    $(if ($status_2_1 -eq 'PASS') {'SPF records present on all domains.'} else {"$($domainsNoSpf.Count) domain(s) missing SPF."})
Write-Check "2.1 SPF records" $status_2_1

# - 2.2 DMARC records -----------------------------------------------------
# Same reasoning as 2.1 above - pure DNS, no Exchange access needed.
$domainsNoDmarc = @($CheckableDomains | Where-Object {
    $d = $_
    try { -not (Resolve-DnsName -Name "_dmarc.$($d.id)" -Type TXT -ErrorAction Stop |
          Where-Object { $_.Text -match 'v=DMARC1' }) } catch { $true }
})
$status_2_2 = if ($domainsNoDmarc.Count -eq 0) {'PASS'} else {'FAIL'}
Add-Result '2-Exchange' '2.2' 'DMARC records shall be configured for all domains' $status_2_2 'high' `
    $(if ($status_2_2 -eq 'PASS') {'DMARC records present.'} else {"$($domainsNoDmarc.Count) domain(s) missing DMARC."})
Write-Check "2.2 DMARC records" $status_2_2

if (-not $SkipExchange) {

    # - 2.3 Anti-phishing (impersonation protection) ---------------------------
    $defaultAntiPhish     = @($AntiPhishPolicies | Where-Object { $_.IsDefault })
    $hasImpersonationProt = @($AntiPhishPolicies | Where-Object {
        $_.EnableTargetedUserProtection -eq $true -or
        $_.EnableOrganizationDomainsProtection -eq $true
    })
    $status_2_3 = if ($hasImpersonationProt) {'PASS'} else {'FAIL'}
    Add-Result '2-Exchange' '2.3' 'Anti-phishing policies with impersonation protection are configured' $status_2_3 'high' `
        $(if ($status_2_3 -eq 'PASS') {'Impersonation protection configured.'} else {'No anti-phishing impersonation policy found.'})
    Write-Check "2.3 Anti-phishing impersonation protection" $status_2_3

    # - 2.4 External email warning --------------------------------------------
    $externalTag = $null
    try { $externalTag = Get-ExternalInOutlook -ErrorAction Stop } catch {}
    $status_2_4  = if ($externalTag -and (Get-Prop $externalTag 'Enabled') -eq $true) {'PASS'} else {'FAIL'}
    Add-Result '2-Exchange' '2.4' 'External email warning shall be enabled' $status_2_4 'medium' `
        $(if ($status_2_4 -eq 'PASS') {'External sender tagging is enabled.'} else {'External sender tagging is disabled.'})
    Write-Check "2.4 External email warning" $status_2_4

    # - 2.5 Suspicious forwarding rules ---------------------------------------
    # Use PSObject.Properties to safely test each property - StrictMode throws on missing properties
    $forwardRules = @($TransportRules | Where-Object {
        $_ -ne $null -and
        ((Get-Prop $_ 'RedirectMessageTo') -or (Get-Prop $_ 'BlindCopyTo'))
    })
    $forwardInbox = @($InboxRules | Where-Object {
        $_ -ne $null -and
        ((Get-Prop $_ 'ForwardTo') -or (Get-Prop $_ 'ForwardAsAttachmentTo') -or (Get-Prop $_ 'RedirectTo'))
    })
    Add-Result '2-Exchange' '2.5' 'Suspicious mail forwarding rules are reviewed' 'MANUAL' 'high' `
        "Transport rules with forwarding: $($forwardRules.Count). Inbox forwarding rules: $($forwardInbox.Count)."
    Write-Check "2.5 Forwarding rules review" 'MANUAL' "$($forwardRules.Count) transport + $($forwardInbox.Count) inbox forward rules"

    # - 2.6 Audit logging enabled ---------------------------------------------
    $orgConfig    = $null
    try { $orgConfig = Get-OrganizationConfig -ErrorAction Stop } catch {}
    $auditEnabled = $orgConfig -and (Get-Prop $orgConfig 'AuditDisabled') -eq $false
    $status_2_6  = if ($auditEnabled) {'PASS'} else {'FAIL'}
    Add-Result '2-Exchange' '2.6' 'Audit logging shall be enabled in Exchange' $status_2_6 'high' `
        $(if ($auditEnabled) {'Exchange audit logging is enabled.'} else {'Exchange audit logging is disabled.'})
    Write-Check "2.6 Exchange audit logging" $status_2_6

} else {
    Write-Host "  Exchange checks skipped (Advanced tier only)" -ForegroundColor Yellow
    # 2.1/2.2 (SPF/DMARC) already ran unconditionally above - only the
    # checks that genuinely need certificate-based Exchange access are
    # tier-gated here.
    foreach ($c in @('2.3','2.4','2.5','2.6')) {
        Add-Result '2-Exchange' $c 'Not included in this assessment tier' 'MANUAL' 'medium' `
            'Deep Exchange Online mailbox security checks require certificate-based access, included in our Advanced assessment. Ask us about upgrading for full mailbox security coverage.'
    }
}

#endregion

#region -- SECTION 3: Teams --------------------------------------------------
Write-Section "3 - Microsoft Teams"

# Teams messaging and meeting policies require PowerShell if Teams module is available
$HasTeamsModule = $false  # GRAPH-ONLY MODE: Teams PowerShell needs a certificate; skipped.
if ($HasTeamsModule) {
    try {
        Connect-MicrosoftTeams -CertificateThumbprint $null `
            -ApplicationId $ClientId -TenantId $TenantId
        $TeamsTenantConfig        = Get-CsTenantFederationConfiguration
        $TeamsMeetingPolicyDefault = Get-CsTeamsMeetingPolicy -Identity Global
        $TeamsMessagingPolicy     = Get-CsTeamsMessagingPolicy -Identity Global
        $TeamsClientConfig        = Get-CsTeamsClientConfiguration -Identity Global
    } catch { $HasTeamsModule = $false }
}

# - 3.1 External user access restricted ---------------------------------------
if ($HasTeamsModule) {
    $extDomainsRestricted = $TeamsTenantConfig.AllowedDomains.Count -gt 0 -or
                            $TeamsTenantConfig.BlockedDomains.Count -gt 0
    $status_3_1 = if ($extDomainsRestricted) {'PASS'} else {'FAIL'}
} else { $status_3_1 = 'MANUAL' }
Add-Result '3-Teams' '3.1' 'Ensure external domains are restricted in the Teams admin center' $status_3_1 'medium' `
    $(if ($status_3_1 -eq 'PASS') {'Teams policy configured.'} else {'Not included in this assessment tier - requires certificate-based Teams PowerShell access, available with our Advanced assessment.'})
Write-Check "3.1 External user access" $status_3_1

# - 3.2 External participants cannot request screen control --------------------
if ($HasTeamsModule) {
    $extControl  = $TeamsMeetingPolicyDefault.AllowExternalParticipantGiveRequestControl
    $status_3_2  = if ($extControl -eq $false) {'PASS'} else {'FAIL'}
} else { $status_3_2 = 'MANUAL' }
Add-Result '3-Teams' '3.2' "Ensure external participants can't give or request control" $status_3_2 'low' `
    $(if ($status_3_2 -eq 'PASS') {'Teams Policy configured accurately.'} else {'Not included in this assessment tier - requires certificate-based Teams PowerShell access, available with our Advanced assessment.'})
Write-Check "3.2 External screen control" $status_3_2

# - 3.3 Anonymous users cannot start meetings ----------------------------------
if ($HasTeamsModule) {
    $anonStart  = $TeamsMeetingPolicyDefault.AllowAnonymousUsersToStartMeeting
    $status_3_3 = if ($anonStart -eq $false) {'PASS'} else {'FAIL'}
} else { $status_3_3 = 'MANUAL' }
Add-Result '3-Teams' '3.3' "Ensure anonymous users and dial-in callers can't start a meeting" $status_3_3 'medium' `
    $(if ($status_3_3 -eq 'PASS') {'Teams Policy configured accurately.'} else {'Not included in this assessment tier - requires certificate-based Teams PowerShell access, available with our Advanced assessment.'})
Write-Check "3.3 Anonymous meeting start" $status_3_3

# - 3.4 Lobby bypass restricted ------------------------------------------------
if ($HasTeamsModule) {
    $lobbyBypass = $TeamsMeetingPolicyDefault.AutoAdmittedUsers
    $status_3_4  = if ($lobbyBypass -in @('OrganizerOnly','InvitedUsers','OrgOnly')) {'PASS'} else {'FAIL'}
} else { $status_3_4 = 'MANUAL' }
Add-Result '3-Teams' '3.4' 'Ensure only people in my org can bypass the lobby' $status_3_4 'low' `
    $(if ($status_3_4 -eq 'PASS') {'Teams Policy configured accurately.'} else {'Not included in this assessment tier - requires certificate-based Teams PowerShell access, available with our Advanced assessment.'})
Write-Check "3.4 Lobby bypass" $status_3_4

# - 3.5 Unmanaged users cannot initiate contact --------------------------------
if ($HasTeamsModule) {
    $unmanagedContact = $TeamsTenantConfig.AllowTeamsConsumer
    $status_3_5       = if ($unmanagedContact -eq $false) {'PASS'} else {'FAIL'}
} else { $status_3_5 = 'MANUAL' }
Add-Result '3-Teams' '3.5' 'Unmanaged users SHALL NOT be enabled to initiate contact with internal users.' $status_3_5 'low' `
    $(if ($status_3_5 -eq 'PASS') {'Teams Policy configured accurately.'} else {'Not included in this assessment tier - requires certificate-based Teams PowerShell access, available with our Advanced assessment.'})
Write-Check "3.5 Unmanaged user contact" $status_3_5

# - 3.6 Skype communication blocked --------------------------------------------
if ($HasTeamsModule) {
    $skypeComm  = $TeamsTenantConfig.AllowPublicUsers
    $status_3_6 = if ($skypeComm -eq $false) {'PASS'} else {'FAIL'}
} else { $status_3_6 = 'MANUAL' }
Add-Result '3-Teams' '3.6' 'Ensure communication with Skype users is disabled' $status_3_6 'low' `
    $(if ($status_3_6 -eq 'PASS') {'Teams Policy configured accurately.'} else {'Not included in this assessment tier - requires certificate-based Teams PowerShell access, available with our Advanced assessment.'})
Write-Check "3.6 Skype communication blocked" $status_3_6

# - 3.7 3rd party file sharing blocked ----------------------------------------
if ($HasTeamsModule) {
    $thirdPartyFiles = $TeamsClientConfig.AllowDropBox -or
                       $TeamsClientConfig.AllowBox -or
                       $TeamsClientConfig.AllowGoogleDrive -or
                       $TeamsClientConfig.AllowEgnyte
    $status_3_7  = if (-not $thirdPartyFiles) {'PASS'} else {'FAIL'}
} else { $status_3_7 = 'MANUAL' }
Add-Result '3-Teams' '3.7' 'Ensure external file sharing in Teams is enabled for only approved cloud storage services' $status_3_7 'low' `
    $(if ($status_3_7 -eq 'PASS') {'3rd party file sharing is blocked.'} else {'Not included in this assessment tier - requires certificate-based Teams PowerShell access, available with our Advanced assessment.'})
Write-Check "3.7 3rd party file sharing" $status_3_7

#endregion

#region -- SECTION 4: Intune --------------------------------------------------
Write-Section "4 - Intune / Endpoint Management"

# - 4.1 Windows Update Rings --------------------------------------------------
$windowsUpdateRings = $UpdateRings
$needsPatchDevices  = @($IntuneDevices | Where-Object {
    $_.operatingSystem -eq 'Windows' -and $_.complianceState -eq 'noncompliant'
})
$status_4_1a = if ($windowsUpdateRings.Count -gt 0) {'PASS'} else {'FAIL'}
Add-Result '4-Intune' '4.1' 'Windows Update Rings shall be configured for Windows Devices' $status_4_1a 'critical' `
    "$($windowsUpdateRings.Count) policies found in Intune. $($needsPatchDevices.Count) devices that need patching."
Write-Check "4.1 Windows Update Rings" $status_4_1a "$($windowsUpdateRings.Count) ring(s)"

# Apple update policy
$appleUpdatePolicy = @($ConfigProfiles | Where-Object {
    $_.'@odata.type' -match 'iosUpdateConfiguration|macOSUpdateConfiguration'
})
$status_4_1b = if ($appleUpdatePolicy.Count -gt 0) {'PASS'} else {'FAIL'}
Add-Result '4-Intune' '4.1' 'Update Policies shall be configured for Apple Devices' $status_4_1b 'high' `
    $(if ($status_4_1b -eq 'PASS') {'Apple update policy found.'} else {'No update policy found in Intune.'})
Write-Check "4.1 Apple Update Policies" $status_4_1b

# - 4.2 Managed devices enrolled in MDM (MANUAL) ------------------------------
Add-Result '4-Intune' '4.2' 'Managed Devices are enrolled in MDM' 'MANUAL' 'medium' 'Manual'
Write-Check "4.2 MDM enrollment" 'MANUAL'

# - 4.3 Personal devices restricted -------------------------------------------
$enrollRestrictions = @()
try {
    $enrollRestrictions = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceEnrollmentConfigurations" `
        -Headers $GraphHeaders
} catch { <# permissions missing - treat as not configured #> }
$personalRestriction = @($enrollRestrictions | Where-Object {
    $_.'@odata.type' -match 'deviceEnrollmentPlatformRestrictionConfiguration' -and
    ((Get-Prop $_ 'windowsRestriction.personalDeviceEnrollmentBlocked') -eq $true -or
     (Get-Prop $_ 'iosRestriction.personalDeviceEnrollmentBlocked')     -eq $true)
})
$status_4_3 = if ($personalRestriction.Count -gt 0) {'PASS'} else {'FAIL'}
Add-Result '4-Intune' '4.3' 'Personal Devices should be restricted from enrolling into the MDM solution' $status_4_3 'low' `
    $(if ($status_4_3 -eq 'PASS') {'Personal device restriction policy found.'} else {'Personal devices are not restricted from enrolling into Intune MDM'})
Write-Check "4.3 Personal device restriction" $status_4_3

# - 4.4 Security Baselines -----------------------------------------------------
$secBaselinePolicy = @($SecurityBaselines | Where-Object {
    $_.displayName -match 'baseline|Baseline'
})
$status_4_4 = if ($secBaselinePolicy.Count -gt 0) {'PASS'} else {'FAIL'}
Add-Result '4-Intune' '4.4' 'Security Baselines should be configured for Windows Devices' $status_4_4 'high' `
    $(if ($status_4_4 -eq 'PASS') {'Security Baseline policy configured in Intune.'} else {'No security baseline found.'})
Write-Check "4.4 Security Baselines" $status_4_4

# - 4.5 Compliance policies for all platforms ----------------------------------
$osList = @('Windows','iOS','Android','macOS')
$coveredOS = $CompliancePolicies | ForEach-Object {
    switch ($_.'@odata.type') {
        '#microsoft.graph.windows10CompliancePolicy' {'Windows'}
        '#microsoft.graph.iosCompliancePolicy'       {'iOS'}
        '#microsoft.graph.androidCompliancePolicy'   {'Android'}
        '#microsoft.graph.macOSCompliancePolicy'     {'macOS'}
    }
} | Sort-Object -Unique
$missingOS  = @($osList | Where-Object { $_ -notin $coveredOS })
$status_4_5 = if ($missingOS.Count -eq 0) {'PASS'} else {'FAIL'}
Add-Result '4-Intune' '4.5' 'Devices compliance policies shall be configured for every supported device platform' $status_4_5 'high' `
    $(if ($status_4_5 -eq 'PASS') {'All platforms are covered by compliance policies.'} else {"Missing compliance policy for: $($missingOS -join ', ')"})
Write-Check "4.5 Compliance policies all platforms" $status_4_5

# - 4.6 Drive encryption required ---------------------------------------------
$notEncryptedCount = $NotEncryptedDevices.Count
$status_4_6 = if ($notEncryptedCount -eq 0) {'PASS'} else {'FAIL'}
Add-Result '4-Intune' '4.6' 'Encryption shall be required on all devices' $status_4_6 'high' `
    "$notEncryptedCount were found that are not encrypted."
Write-Check "4.6 Device encryption" $status_4_6 "$notEncryptedCount unencrypted"

# - 4.7 Lockout / password policy ---------------------------------------------
$lockoutPolicy = @($CompliancePolicies | Where-Object {
    $_ -ne $null -and
    ((Get-Prop $_ 'passwordRequired') -eq $true -or (Get-Prop $_ 'passcodeRequired') -eq $true)
})
$status_4_7 = if ($lockoutPolicy.Count -gt 0) {'PASS'} else {'FAIL'}
Add-Result '4-Intune' '4.7' 'Lockout screen and password settings shall be configured for each device' $status_4_7 'medium' `
    $(if ($status_4_7 -eq 'PASS') {'Password compliance policy found.'} else {'No device lockout policy configured in Intune.'})
Write-Check "4.7 Device lockout / password" $status_4_7

# - 4.8 App protection - mobile -----------------------------------------------
$hasAppProt = ($AppProtectioniOS.Count -gt 0 -or $AppProtectionAndroid.Count -gt 0)
$status_4_8 = if ($hasAppProt) {'PASS'} else {'FAIL'}
Add-Result '4-Intune' '4.8' 'App Protection policies should be created for mobile devices' $status_4_8 'medium' `
    $(if ($hasAppProt) {"Polices are in Intune for iOS=$($AppProtectioniOS.Count) and Android=$($AppProtectionAndroid.Count)"} else {'No app protection policies found.'})
Write-Check "4.8 App protection policies" $status_4_8

# - 4.9 App deployment through Intune -----------------------------------------
$managedApps = @()
try {
    $managedApps = Invoke-GraphAll `
        -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps?`$top=999&`$select=id,displayName,isAssigned" `
        -Headers $GraphHeaders
} catch { <# permissions missing - skip #> }
$assignedApps = @($managedApps | Where-Object { $_.isAssigned -eq $true })
$status_4_9   = if ($assignedApps.Count -gt 0) {'PASS'} else {'FAIL'}
Add-Result '4-Intune' '4.9' 'Authorized Applications should be deployed to managed devices' $status_4_9 'medium' `
    $(if ($status_4_9 -eq 'PASS') {'Applications are being deployed through Intune'} else {'No assigned apps found in Intune.'})
Write-Check "4.9 Intune app deployment" $status_4_9

# - 4.10 LAPS ------------------------------------------------------------------
$lapsEnabled = $null
try {
    $lapsEnabled = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations?`$filter=startswith(displayName,'LAPS')" `
        -Headers $GraphHeaders -ErrorAction Stop
} catch { <# permissions or endpoint missing - skip #> }
$hasLaps    = ($lapsEnabled  -and @($lapsEnabled.value).Count  -gt 0) -or
              ($LapsPolicy.Count -gt 0)
$status_4_10 = if ($hasLaps) {'PASS'} else {'FAIL'}
Add-Result '4-Intune' '4.10' 'Ensure Local Administrator Password Solution is enabled' $status_4_10 'high' `
    $(if ($hasLaps) {'Local admin password solution is enabled'} else {'LAPS not configured.'})
Write-Check "4.10 LAPS" $status_4_10

#endregion

#region -- SECTION 5: SharePoint & OneDrive -----------------------------------
Write-Section "5 - SharePoint & OneDrive"

# - 5.1 External sharing restricted -------------------------------------------
$sharingLevel   = $SPSettings.sharingCapability
# Values: disabled | existingExternalUserSharingOnly | externalUserSharingOnly | externalUserAndGuestSharing
$sharingOk      = $sharingLevel -in @('disabled','existingExternalUserSharingOnly')
$status_5_1a    = if ($sharingOk) {'PASS'} else {'FAIL'}
Add-Result '5-SharePoint' '5.1' 'Ensure SharePoint external sharing is managed through domain whitelist/blacklists' $status_5_1a 'medium' `
    $(if ($status_5_1a -eq 'PASS') {'SharePoint sharing restricted.'} else {'SharePoint Settings Misconfigured'})
Write-Check "5.1 SharePoint external sharing" $status_5_1a "Sharing level: $sharingLevel"

Add-Result '5-SharePoint' '5.1' 'Ensure link sharing is restricted in SharePoint and OneDrive' 'MANUAL' 'medium' 'Manual.'
Write-Check "5.1 Link sharing restriction" 'MANUAL'

# - 5.2 Anyone link expiration ------------------------------------------------
Add-Result '5-SharePoint' '5.2' 'Expiration Date SHOULD Be Set for Anyone Links' 'MANUAL' 'low' 'Manual.'
Write-Check "5.2 Anyone link expiration" 'MANUAL'

# Collect top sites with public/anonymous links for data tab.
# Anonymous ("Anyone") links are sharing links whose link.scope is 'anonymous'
# (or 'anyone'); they carry NO granted user identity, so we must NOT require one.
$SPSitesWithPublicLinks = foreach ($site in ($SPSites | Select-Object -First 100)) {
    try {
        $sitePerms = Invoke-GraphAll `
            -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/permissions" `
            -Headers $GraphHeaders
        $publicCount = (@($sitePerms | Where-Object {
            $scope = "$(Get-Prop $_ 'link.scope')".ToLower()
            $scope -eq 'anonymous' -or $scope -eq 'anyone'
        })).Count
        if ($publicCount -gt 0) {
            [PSCustomObject]@{
                SiteDisplayName = $site.displayName
                WebUrl          = $site.webUrl
                PublicLinkCount = $publicCount
            }
        }
    } catch {}
}
$SPSitesWithPublicLinks = @($SPSitesWithPublicLinks)

#endregion

#region -- SECTION 6: Defender -----------------------------------------------
Write-Section "6 - Microsoft Defender"

# - 6.1 Security awareness training (MANUAL) ----------------------------------
Add-Result '6-Defender' '6.1' 'Attack simulations shall be periodically conducted' 'MANUAL' 'medium' 'Manual'
Write-Check "6.1 Security awareness training" 'MANUAL'

# - 6.2 Defender AV via Intune ------------------------------------------------
$defenderAvPolicy = @($ConfigProfiles | Where-Object {
    $_.'@odata.type' -match 'windowsDefender|antivirus|endpointProtection'
})
$status_6_2 = if ($defenderAvPolicy.Count -gt 0) {'PASS'} else {'FAIL'}
Add-Result '6-Defender' '6.2' 'Microsoft Defender Antivirus is deployed and managed through Microsoft Intune' $status_6_2 'high' `
    $(if ($status_6_2 -eq 'PASS') {'Microsoft Defender Antivirus policy configured in Intune.'} else {'No Defender AV policy found in Intune.'})
Write-Check "6.2 Defender AV via Intune" $status_6_2

# - 6.3 Defender for Endpoint enrollment --------------------------------------
# Check via security/microsoft.graph.security.deviceInventories or use managedDevices
$mdeEnrolledCount = @($IntuneDevices | Where-Object {
    $_.operatingSystem -eq 'Windows'
}).Count
$status_6_3 = if ($mdeEnrolledCount -gt 0) {'PASS'} else {'FAIL'}
Add-Result '6-Defender' '6.3' 'Devices shall be enrolled for Defender for Business or Defender for Endpoint' $status_6_3 'high' `
    "$mdeEnrolledCount Devices found in Defender via Intune"
Write-Check "6.3 MDE enrollment" $status_6_3 "$mdeEnrolledCount devices"

# - 6.4 Firewall policy --------------------------------------------------------
$firewallPolicy = @($ConfigProfiles | Where-Object {
    $_.'@odata.type' -match 'windowsFirewall' -or
    $_.displayName -match 'Firewall'
})
$status_6_4 = if ($firewallPolicy.Count -gt 0) {'PASS'} else {'FAIL'}
Add-Result '6-Defender' '6.4' 'Firewall Policies are configured for Windows Devices' $status_6_4 'high' `
    $(if ($status_6_4 -eq 'PASS') {'Firewall policy found.'} else {'No Windows firewall policy found in Intune.'})
Write-Check "6.4 Firewall policy" $status_6_4

# - 6.5 Safe Links ------------------------------------------------------------
$hasActiveSafeLinks = $false
if (-not $SkipExchange -and $SafeLinksPolicies) {
    $hasActiveSafeLinks = (@($SafeLinksPolicies | Where-Object {
        $_ -ne $null -and (Get-Prop $_ 'IsEnabled') -ne $false
    })).Count -gt 0
}
$status_6_5 = if ($hasActiveSafeLinks) {'PASS'} elseif ($SkipExchange) {'MANUAL'} else {'FAIL'}
Add-Result '6-Defender' '6.5' 'Safe Links policies are configured' $status_6_5 'high' `
    $(if ($status_6_5 -eq 'PASS') {'Safe Links policies are present and correctly configured.'} `
      elseif ($status_6_5 -eq 'MANUAL') {'Not included in this assessment tier - requires certificate-based Exchange Online access, available with our Advanced assessment.'} `
      else {'No active Safe Links policy found.'})
Write-Check "6.5 Safe Links" $status_6_5

# - 6.6 Safe Attachments ------------------------------------------------------
$hasActiveSafeAtt = $false
if (-not $SkipExchange -and $SafeAttachmentPolicies) {
    $hasActiveSafeAtt = (@($SafeAttachmentPolicies | Where-Object {
        $_ -ne $null -and (Get-Prop $_ 'Enable') -eq $true
    })).Count -gt 0
}
$status_6_6 = if ($hasActiveSafeAtt) {'PASS'} elseif ($SkipExchange) {'MANUAL'} else {'FAIL'}
Add-Result '6-Defender' '6.6' 'Safe Attachment Policies are configured' $status_6_6 'high' `
    $(if ($status_6_6 -eq 'PASS') {'Active Safe Attachment policy found'} `
      elseif ($status_6_6 -eq 'MANUAL') {'Not included in this assessment tier - requires certificate-based Exchange Online access, available with our Advanced assessment.'} `
      else {'No active Safe Attachment policy.'})
Write-Check "6.6 Safe Attachments" $status_6_6

# - 6.7 Tamper Protection (MANUAL / Defender for Endpoint API) ----------------
Add-Result '6-Defender' '6.7' 'Turn on Tamper Protection' 'MANUAL' 'high' 'Manual'
Write-Check "6.7 Tamper Protection" 'MANUAL'

# - 6.8 Attack Surface Reduction rules ----------------------------------------
$asrPolicy = @($ConfigProfiles | Where-Object {
    $_.'@odata.type' -match 'windowsDefenderAdvancedThreatProtection' -or
    $_.displayName -match 'Attack Surface|ASR'
})
$status_6_8 = if ($asrPolicy.Count -gt 0) {'PASS'} else {'FAIL'}
Add-Result '6-Defender' '6.8' 'Attack Surface Reduction rules shall be configured' $status_6_8 'high' `
    $(if ($status_6_8 -eq 'PASS') {'Attack surface reduction policy found in Intune'} else {'No ASR policy found.'})
Write-Check "6.8 Attack Surface Reduction" $status_6_8

# - 6.9 Defender for Cloud Apps ------------------------------------------------
$defenderCloudApps = @($SecureScoreControls | Where-Object {
    $_ -ne $null -and (Get-Prop $_ 'controlName') -match 'CloudApp|MCAS'
})
$hasMCASLicense = (@($Org.assignedPlans | Where-Object {
    $_ -ne $null -and
    (Get-Prop $_ 'service') -match 'MicrosoftDefenderCloudApps' -and
    (Get-Prop $_ 'capabilityStatus') -eq 'Enabled'
})).Count -gt 0

$status_6_9 = if ($hasMCASLicense) {'PASS'} else {'FAIL'}
Add-Result '6-Defender' '6.9' 'Cloud App Discovery is configured and apps are periodically reviewed' $status_6_9 'medium' `
    $(if ($hasMCASLicense) {'Your tenant is licensed for Defender for Cloud Apps'} else {'Defender for Cloud Apps license not found.'})
Write-Check "6.9 Defender for Cloud Apps" $status_6_9

#endregion

#region -- SECTION 7: Purview -------------------------------------------------
Write-Section "7 - Purview / Compliance"

# - 7.1 Backups (MANUAL) ------------------------------------------------------
Add-Result '7-Purview' '7.1' 'Maintain 3rd party backups of Microsoft 365 data' 'MANUAL' 'medium' 'Manual'
Write-Check "7.1 Third-party backups" 'MANUAL'

# - 7.2 Unified Audit Log enabled ---------------------------------------------
$auditStatus = $null
try {
    $auditStatus = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/security/secureScores?`$top=1" `
        -Headers $GraphHeaders -ErrorAction Stop
} catch { <# fall through to Exchange check below #> }
# Try direct check via admin audit log config
$auditLogEnabled = $false
if (-not $SkipExchange) {
    try {
        $adminAudit      = Get-AdminAuditLogConfig
        $auditLogEnabled = $adminAudit.UnifiedAuditLogIngestionEnabled
    } catch { $auditLogEnabled = $false }
}
$status_7_2 = if ($auditLogEnabled) {'PASS'} elseif ($SkipExchange) {'MANUAL'} else {'FAIL'}
Add-Result '7-Purview' '7.2' 'Audit Logging SHALL Be Enabled' $status_7_2 'high' `
    $(if ($auditLogEnabled) {'Audit Logging Enabled.'} else {'Audit logging check inconclusive - verify manually.'})
Write-Check "7.2 Unified Audit Log" $status_7_2

# - 7.3 Retention policies (MANUAL) -------------------------------------------
Add-Result '7-Purview' '7.3' 'Retention Policies shall be configured' 'MANUAL' 'medium' 'Manual'
Write-Check "7.3 Retention policies" 'MANUAL'

# - 7.4 Sensitivity Labels ----------------------------------------------------
$sensitivityCtrl = $ControlMap['SensitivityLabels']
$status_7_4 = if ($sensitivityCtrl -and (Get-Prop $sensitivityCtrl 'implementationStatus') -eq 'Implemented') {'PASS'} else {'FAIL'}
Add-Result '7-Purview' '7.4' 'Information Protection Labels shall be configured' $status_7_4 'high' `
    $(if ($status_7_4 -eq 'PASS') {'Sensitivity labels configured.'} else {'Secure Score Controls.'})
Write-Check "7.4 Sensitivity Labels" $status_7_4

# - 7.5 DLP Policies ----------------------------------------------------------
$dlpCtrl = $ControlMap['DLPPolicy']
$status_7_5 = if ($dlpCtrl -and (Get-Prop $dlpCtrl 'implementationStatus') -eq 'Implemented') {'PASS'} else {'FAIL'}
Add-Result '7-Purview' '7.5' 'Data loss prevention policies shall be configured' $status_7_5 'high' `
    $(if ($status_7_5 -eq 'PASS') {'Secure Score Controls.'} else {'No DLP policies found.'})
Write-Check "7.5 DLP Policies" $status_7_5

#endregion

#region -- Summary Metrics ----------------------------------------------------
Write-Section "Assessment Summary"

$Total   = $Results.Count
$Passed  = @($Results | Where-Object { $_.Status -eq 'PASS'   }).Count
$Failed  = @($Results | Where-Object { $_.Status -eq 'FAIL'   }).Count
$Manual  = @($Results | Where-Object { $_.Status -eq 'MANUAL' }).Count
$Warned  = @($Results | Where-Object { $_.Status -eq 'WARN'   }).Count

$PassPct = [math]::Round(($Passed / $Total) * 100, 0)

Write-Host ""
Write-Host "  Secure Score       : $($SecureScore.currentScore)/$($SecureScore.maxScore) ($SecureScorePct%)" -ForegroundColor Cyan
Write-Host "  Total users        : $($AllUsers.Count)"            -ForegroundColor White
Write-Host "  Global Admins      : $gaCount"                       -ForegroundColor $(if ($gaCount -ge 2 -and $gaCount -le 4) {'Green'} else {'Red'})
Write-Host "  Users without MFA  : $noMfaCount"                   -ForegroundColor $(if ($noMfaCount -eq 0) {'Green'} else {'Red'})
Write-Host "  Users with weak MFA: $($WeakMfaUsers.Count)"        -ForegroundColor $(if ($WeakMfaUsers.Count -eq 0) {'Green'} else {'Yellow'})
Write-Host "  Risky users        : $($RiskyUsers.Count)"          -ForegroundColor $(if ($RiskyUsers.Count -eq 0) {'Green'} else {'Red'})
Write-Host "  Entra devices      : $($EntraDevices.Count)"        -ForegroundColor White
Write-Host "  Stale devices      : $($StaleDevices.Count)"        -ForegroundColor $(if ($StaleDevices.Count -eq 0) {'Green'} else {'Yellow'})
Write-Host "  Non-compliant dev  : $($NonCompliantDevices.Count)" -ForegroundColor $(if ($NonCompliantDevices.Count -eq 0) {'Green'} else {'Red'})
Write-Host "  Unencrypted dev    : $notEncryptedCount"            -ForegroundColor $(if ($notEncryptedCount -eq 0) {'Green'} else {'Red'})
Write-Host "  Dormant accounts   : $dormantCount"                 -ForegroundColor $(if ($dormantCount -eq 0) {'Green'} else {'Yellow'})
Write-Host "  Enterprise apps    : $appCount"                     -ForegroundColor White
Write-Host ""
Write-Host "  Controls: PASS=$Passed  FAIL=$Failed  MANUAL=$Manual  WARN=$Warned  TOTAL=$Total" -ForegroundColor White
Write-Host "  Automated pass rate: $PassPct%" -ForegroundColor Cyan

#endregion

#region -- Export Results -----------------------------------------------------
Write-Section "Exporting results"

$Timestamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$JsonPath    = Join-Path $OutputFolder "Assessment_Results_$Timestamp.json"
$SummaryPath = Join-Path $OutputFolder "Assessment_Summary_$Timestamp.csv"

# -- Email Health (per-domain SPF/DKIM/DMARC + mail-flow) ---------------------
# Ensure Exchange-only variables exist when -SkipExchange was used
if (-not (Get-Variable -Name DkimConfigs -Scope Script -ErrorAction SilentlyContinue) -and
    -not (Get-Variable -Name DkimConfigs -ErrorAction SilentlyContinue)) { $DkimConfigs = @() }
if (-not (Get-Variable -Name mfReport -ErrorAction SilentlyContinue)) { $mfReport = @() }

Write-Host "  Computing Email Health per domain..." -NoNewline
$EmailHealth = @()
foreach ($dom in $Domains) {
    $dn        = Get-Prop $dom 'id'
    if ([string]::IsNullOrWhiteSpace($dn)) { continue }
    $isDefault = [bool](Get-Prop $dom 'isDefault')
    $isOnMs    = $dn -match '\.onmicrosoft\.com$'

    if ($isOnMs) {
        # Microsoft routing domains - DNS records are managed by Microsoft
        $spf = 'NA'; $verified = 'NA'; $dkim = 'NA'; $dmarc = 'NA'; $dmarcPolicy = ''
    } else {
        # SPF
        $spf = 'FAIL'
        try {
            if (Resolve-DnsName -Name $dn -Type TXT -ErrorAction Stop |
                Where-Object { $_.Text -match 'v=spf1' }) { $spf = 'PASS' }
        } catch { $spf = 'FAIL' }

        # Verified (has MX records)
        $verified = 'FAIL'
        try {
            if (Resolve-DnsName -Name $dn -Type MX -ErrorAction Stop) { $verified = 'PASS' }
        } catch { $verified = 'FAIL' }

        # DKIM - prefer Exchange config, fall back to DNS CNAME selector1
        $dkim = 'FAIL'
        $dkimCfg = $DkimConfigs | Where-Object { (Get-Prop $_ 'Domain') -eq $dn } | Select-Object -First 1
        if ($dkimCfg -and [bool](Get-Prop $dkimCfg 'Enabled')) {
            $dkim = 'PASS'
        } else {
            try {
                if (Resolve-DnsName -Name "selector1._domainkey.$dn" -Type CNAME -ErrorAction Stop) { $dkim = 'PASS' }
            } catch { $dkim = 'FAIL' }
        }

        # DMARC + policy
        $dmarc = 'FAIL'; $dmarcPolicy = 'Missing'
        try {
            $dmarcRec = Resolve-DnsName -Name "_dmarc.$dn" -Type TXT -ErrorAction Stop |
                        Where-Object { $_.Text -match 'v=DMARC1' } | Select-Object -First 1
            if ($dmarcRec) {
                $dmarcText = ($dmarcRec.Text -join '')
                if     ($dmarcText -match 'p\s*=\s*reject')     { $dmarcPolicy = 'Reject';     $dmarc = 'PASS' }
                elseif ($dmarcText -match 'p\s*=\s*quarantine') { $dmarcPolicy = 'Quarantine'; $dmarc = 'PASS' }
                elseif ($dmarcText -match 'p\s*=\s*none')       { $dmarcPolicy = 'None';       $dmarc = 'FAIL' }
                else                                            { $dmarcPolicy = 'None';       $dmarc = 'FAIL' }
            }
        } catch { $dmarc = 'FAIL'; $dmarcPolicy = 'Missing' }
    }

    $EmailHealth += [PSCustomObject]@{
        Domain        = $dn
        IsDefault     = $isDefault
        IsOnMicrosoft = $isOnMs
        Spf           = $spf
        Verified      = $verified
        Dkim          = $dkim
        Dmarc         = $dmarc
        DmarcPolicy   = $dmarcPolicy
    }
}
Write-Host " $($EmailHealth.Count) domains" -ForegroundColor Green

# -- Mail-flow summary (last 30 days) -----------------------------------------
$mfScanned = 0; $mfDelivered = 0; $mfBlocked = 0

# Try Exchange-based stats first (available when -SkipExchange is not set)
foreach ($row in $mfReport) {
    $cnt = 0
    $rawCnt = Get-Prop $row 'MessageCount'
    if ($rawCnt) { [int]::TryParse("$rawCnt", [ref]$cnt) | Out-Null }
    $mfScanned += $cnt
    $evt = "$(Get-Prop $row 'EventType')"
    if ($evt -match 'GoodMail|Delivered') { $mfDelivered += $cnt } else { $mfBlocked += $cnt }
}

# Fall back to Graph Reports API data when Exchange was skipped
if ($mfScanned -eq 0 -and $MailActivityReport.Count -gt 0) {
    foreach ($day in $MailActivityReport) {
        $s = 0; $r = 0; $sp = 0
        try { $s  = [int64]($day.'Send Count') }       catch {}
        try { $r  = [int64]($day.'Receive Count') }    catch {}
        try { $sp = [int64]($day.'Spam Receive Count') } catch {}
        $mfDelivered += ($s + $r)
        $mfBlocked   += $sp
    }
    $mfScanned = $mfDelivered + $mfBlocked
}

$MailFlow = @{
    EmailsScanned   = $mfScanned
    EmailsDelivered = $mfDelivered
    EmailsBlocked   = $mfBlocked
}

# -- JSON (full results) ------------------------------------------------------
$exportObj = [ordered]@{
    GeneratedAt    = (Get-Date -Format 'o')
    # Read by New-M365Report.ps1 to select tier-accurate report copy (this
    # is the Graph-only, client-secret Essential-tier script - see the
    # SYNOPSIS above).
    AssessmentTier = 'Essential'
    TenantId       = $TenantId
    DefaultDomain  = $DefaultDomain
    SecureScore    = @{
        Current    = $SecureScore.currentScore
        Max        = $SecureScore.maxScore
        Percentage = $SecureScorePct
        History    = $SecureScoreHistory | Select-Object -First 30 createdDateTime, currentScore, maxScore
    }
    UserHealth     = @{
        TotalUsers          = $AllUsers.Count
        GlobalAdmins        = $gaCount
        UsersWithoutMFA     = $noMfaCount
        UsersWithWeakMFA    = $WeakMfaUsers.Count
        RiskyUsers          = $RiskyUsers.Count
        DormantUsers        = $dormantCount
    }
    DeviceHealth   = @{
        EntraDevices         = $EntraDevices.Count
        StaleDevices         = $StaleDevices.Count
        NonCompliantDevices  = $NonCompliantDevices.Count
        NotEncryptedDevices  = $notEncryptedCount
    }
    ApplicationsData = @{
        EnterpriseApps      = $appCount
        SPSharingSetting    = Get-Prop $SPSettings 'sharingCapability'
        SPSitesPublicLinks  = $SPSitesWithPublicLinks
        # Full SharePoint site inventory so the report can list ALL sites,
        # not just those with public links.
        SPTotalSites        = $SPSites.Count
        SPSites             = @($SPSites | Select-Object `
                                @{N='DisplayName';E={ $_.displayName }}, `
                                @{N='Name';       E={ $_.name }}, `
                                @{N='WebUrl';     E={ $_.webUrl }})
    }
    EmailHealth    = $EmailHealth
    MailFlow       = $MailFlow
    Licensing      = @($SubscribedSkus | ForEach-Object {
        $consumed = 0; $total = 0
        $cp = $_.PSObject.Properties['consumedUnits']; if ($cp) { $consumed = $cp.Value }
        $pp = $_.PSObject.Properties['prepaidUnits'];  if ($pp) { $total = $pp.Value.enabled }
        [PSCustomObject]@{
            SkuPartNumber = [string](Get-Prop $_ 'skuPartNumber')
            SkuId         = [string](Get-Prop $_ 'skuId')
            DisplayName   = [string](Get-Prop $_ 'skuPartNumber')
            Consumed      = $consumed
            Total         = $total
        }
    })
    AuthMethods    = if ($AuthMethodsPolicy) {
        @($AuthMethodsPolicy.authenticationMethodConfigurations | ForEach-Object {
            [PSCustomObject]@{
                Method = [string](Get-Prop $_ 'id')
                State  = [string](Get-Prop $_ 'state')
            }
        })
    } else { @() }
    GuestAccess    = @{
        GuestUserCount = $GuestUsers.Count
        TotalUsers     = $AllUsers.Count
        GuestInviteSettings = if ($AuthPolicy) {
            [string](Get-Prop $AuthPolicy 'allowInvitesFrom')
        } else { 'Unknown' }
        CrossTenantAccessPolicy = if ($ExternalCollabSettings) {
            [string]($ExternalCollabSettings | ConvertTo-Json -Depth 3 -Compress)
        } else { 'Not available' }
    }
    Controls       = $Results
}
$exportObj | ConvertTo-Json -Depth 10 | Out-File $JsonPath -Encoding UTF8
Write-Host "  JSON  -> $JsonPath" -ForegroundColor Green

# -- CSV summary --------------------------------------------------------------
$Results | Export-Csv $SummaryPath -NoTypeInformation -Encoding UTF8
Write-Host "  CSV   -> $SummaryPath" -ForegroundColor Green

# -- Excel workbook (if ImportExcel module available) -------------------------
if ($HasImportExcel) {
  try {
    $XlsxPath = Join-Path $OutputFolder "Assessment_Workbook_$Timestamp.xlsx"

    # Tab 1 - Control Results
    $Results | Export-Excel $XlsxPath -WorksheetName 'Controls' -AutoFilter -BoldTopRow -AutoSize

    # Tab 2 - User list (with MFA & sign-in details)
    $userSheet = $AllUsers | Select-Object displayName, userPrincipalName, accountEnabled, userType,
        @{N='LastSignIn'; E={
            $sia = $_.PSObject.Properties['signInActivity']
            if ($sia -and $sia.Value) { $sia.Value.lastSignInDateTime } else { $null }
        }},
        @{N='IsMfaRegistered'; E={
            $upn = $_.userPrincipalName
            $u   = $MfaReg | Where-Object { $_.userPrincipalName -eq $upn }
            if ($u) { $u.isMfaRegistered } else { $null }
        }}
    $userSheet | Export-Excel $XlsxPath -WorksheetName 'Users' -AutoFilter -BoldTopRow -AutoSize -Append

    # Tab 3 - Users without MFA
    ($MfaReg | Where-Object { -not $_.isMfaRegistered }) |
        Select-Object userPrincipalName, displayName, isMfaRegistered,
            @{N='methodsRegistered'; E={($_.methodsRegistered -join ', ')}} |
        Export-Excel $XlsxPath -WorksheetName 'UsersNoMFA' -AutoFilter -BoldTopRow -AutoSize -Append

    # Tab 4 - Users with weak MFA
    $WeakMfaUsers |
        Select-Object userPrincipalName, displayName,
            @{N='methodsRegistered'; E={($_.methodsRegistered -join ', ')}} |
        Export-Excel $XlsxPath -WorksheetName 'UsersWeakMFA' -AutoFilter -BoldTopRow -AutoSize -Append

    # Tab 5 - Global Admins
    $GlobalAdmins |
        Select-Object displayName, userPrincipalName, onPremisesSyncEnabled |
        Export-Excel $XlsxPath -WorksheetName 'GlobalAdmins' -AutoFilter -BoldTopRow -AutoSize -Append

    # Tab 6 - Dormant Users
    $DormantUsers |
        Select-Object displayName, userPrincipalName,
            @{N='LastSignIn'; E={
                $sia = $_.PSObject.Properties['signInActivity']
                if ($sia -and $sia.Value) { $sia.Value.lastSignInDateTime } else { 'Never' }
            }} |
        Export-Excel $XlsxPath -WorksheetName 'DormantUsers' -AutoFilter -BoldTopRow -AutoSize -Append

    # Tab 7 - Risky Users
    $RiskyUsers |
        Select-Object userDisplayName, userPrincipalName, riskLevel, riskState, riskDetail |
        Export-Excel $XlsxPath -WorksheetName 'RiskyUsers' -AutoFilter -BoldTopRow -AutoSize -Append

    # Tab 8 - Devices (Intune)
    $IntuneDevices |
        Select-Object deviceName, operatingSystem, osVersion,
            complianceState, isEncrypted, managedDeviceOwnerType, lastSyncDateTime |
        Export-Excel $XlsxPath -WorksheetName 'Devices' -AutoFilter -BoldTopRow -AutoSize -Append

    # Tab 9 - Non-Compliant Devices
    $NonCompliantDevices |
        Select-Object deviceName, operatingSystem, osVersion, complianceState, lastSyncDateTime |
        Export-Excel $XlsxPath -WorksheetName 'NonCompliantDevices' -AutoFilter -BoldTopRow -AutoSize -Append

    # Tab 10 - Devices not encrypted
    $NotEncryptedDevices |
        Select-Object deviceName, operatingSystem, osVersion, isEncrypted, lastSyncDateTime |
        Export-Excel $XlsxPath -WorksheetName 'DevicesNoEncryption' -AutoFilter -BoldTopRow -AutoSize -Append

    # Tab 11 - Conditional Access Policies
    $CAPolicies |
        Select-Object displayName, state,
            @{N='IncludeUsers'; E={ @(Get-Prop $_ 'conditions.users.includeUsers')  -join ',' }},
            @{N='IncludeRoles'; E={ @(Get-Prop $_ 'conditions.users.includeRoles')  -join ',' }},
            @{N='Platforms';    E={ @(Get-Prop $_ 'conditions.platforms.includePlatforms') -join ',' }},
            @{N='GrantControls';E={ @(Get-Prop $_ 'grantControls.builtInControls')  -join ',' }} |
        Export-Excel $XlsxPath -WorksheetName 'ConditionalAccessPolicies' -AutoFilter -BoldTopRow -AutoSize -Append

    # Tab 12 - Enterprise Applications
    $EnterpriseApps |
        Select-Object displayName, appId, publisherName, signInAudience |
        Export-Excel $XlsxPath -WorksheetName 'EnterpriseApps' -AutoFilter -BoldTopRow -AutoSize -Append

    # Tab 13 - SharePoint Settings
    [PSCustomObject]@{
        SharingCapability            = Get-Prop $SPSettings 'sharingCapability'
        DefaultSharingLinkType       = Get-Prop $SPSettings 'defaultSharingLinkType'
        DefaultLinkPermission        = Get-Prop $SPSettings 'defaultLinkPermission'
        AllowedDomainGuidsForSyncApp = (@(Get-Prop $SPSettings 'allowedDomainGuidsForSyncApp') -join ',')
    } | Export-Excel $XlsxPath -WorksheetName 'SharePointSettings' -AutoFilter -BoldTopRow -AutoSize -Append

    # Tab 13b - SharePoint Site Inventory (all discovered sites)
    if ($SPSites.Count -gt 0) {
        $SPSites |
            Select-Object `
                @{N='DisplayName';E={ $_.displayName }}, `
                @{N='Name';       E={ $_.name }}, `
                @{N='WebUrl';     E={ $_.webUrl }} |
            Export-Excel $XlsxPath -WorksheetName 'SharePointSites' -AutoFilter -BoldTopRow -AutoSize -Append
    }

    # Tab 14 - Domains
    $Domains |
        Select-Object id, isVerified, isDefault, authenticationType,
            passwordValidityPeriodInDays, passwordNotificationWindowInDays |
        Export-Excel $XlsxPath -WorksheetName 'Domains' -AutoFilter -BoldTopRow -AutoSize -Append

    # Tab 15 - Secure Score History
    $SecureScoreHistory |
        Select-Object createdDateTime, currentScore, maxScore,
            @{N='Percentage'; E={ [math]::Round(($_.currentScore/$_.maxScore)*100,1) }} |
        Export-Excel $XlsxPath -WorksheetName 'SecureScoreHistory' -AutoFilter -BoldTopRow -AutoSize -Append

    Write-Host "  XLSX  -> $XlsxPath" -ForegroundColor Green
  } catch {
    Write-Host "  XLSX generation failed - continuing without workbook: $($_.Exception.Message)" -ForegroundColor Yellow
  }
} else {
    Write-Host "  (ImportExcel module not available - install with: Install-Module ImportExcel)" -ForegroundColor Yellow
    Write-Host "  CSV written instead: $SummaryPath" -ForegroundColor Yellow
}

#endregion

#region -- Cleanup ------------------------------------------------------------
if (-not $SkipExchange) {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
if ($HasTeamsModule) {
    Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue
}
#endregion

Write-Host "`n  Assessment complete. Review $OutputFolder for deliverables.`n" -ForegroundColor Green
