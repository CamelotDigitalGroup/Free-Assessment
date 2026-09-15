<#
.SYNOPSIS
    Generates a branded Camelot Digital Group Cloud Assessment Report (HTML + PDF)
    from the JSON output produced by M365-SecurityAssessment.ps1.

.DESCRIPTION
    Reads an Assessment_Results_*.json file and produces a self-contained, dark-theme
    HTML report matching the Camelot Digital Group Cloud Assessment Report design,
    then converts it to PDF using Microsoft Edge or Google Chrome in headless mode.

    The generated HTML has no external dependencies (all CSS, SVG and logic inline).

.PARAMETER JsonPath
    Path to the Assessment_Results_*.json file.

.PARAMETER CustomerName
    Customer display name (e.g. "The Stronach Group"). If blank, auto-detected from
    the tenant default domain.

.PARAMETER OutputFolder
    Folder to write the report to. Defaults to the folder containing JsonPath.

.PARAMETER LogoPath
    Optional PNG/SVG logo to embed on the cover page. If blank, the built-in
    Camelot Digital Group shield + wordmark is used.

.PARAMETER PreparedBy
    Name shown as report preparer. Defaults to "Camelot Digital Group".

.PARAMETER TypicalClientScore
    Benchmark score (percentage) for the comparison bar. Defaults to 85.

.EXAMPLE
    .\New-M365Report.ps1 -JsonPath "C:\Assessments\Output\Assessment_Results_20260725_175757.json" `
                         -CustomerName "The Stronach Group" `
                         -OutputFolder "C:\Assessments\Output"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$JsonPath,

    [string]$CustomerName = '',

    [string]$OutputFolder = '',

    [string]$LogoPath = '',

    [string]$PreparedBy = 'Camelot Digital Group',

    [int]$TypicalClientScore = 85
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Palette - matches the live site's actual brand tokens exactly
# (website/tailwind.config.ts "heritage"/"sovereign"/"admiralty" colors),
# not the earlier dark "CDG shield" concept this report used to carry.
# Cream/parchment page background is a deliberate match to the site's own
# use of the same tone for premium, "heritage stationery" surfaces (see
# lib/email.ts's #FAF8F0 email backgrounds) - not a generic light theme.
# ---------------------------------------------------------------------------
# Verified 2026-09-13 against the actual Camelot Brand Guidelines v4.0
# ("The Admiralty Edition") - every value below is one of the six
# specified colors (Admiralty, Signal Teal, Sovereign Gold, Parchment,
# Charcoal, Ice) used exactly, or that same color at reduced opacity for
# a subtler tint - not an invented neutral family. Per the guide:
# "Sovereign Gold should not carry body text - large decorative only."
$COLOR_BG       = '#F6F3EC'   # Parchment - document background
$COLOR_CARD     = '#FFFFFF'   # card surface on the parchment page
$COLOR_BORDER   = 'rgba(7,26,51,0.14)'    # Admiralty at low opacity - subtle dividers
$COLOR_BORDER2  = 'rgba(7,26,51,0.10)'    # Admiralty at lower opacity - table rows
$COLOR_STRIPE   = 'rgba(7,26,51,0.035)'   # Admiralty at very low opacity - zebra striping
$COLOR_TEXT     = '#071A33'   # Admiralty - headings
$COLOR_BODY     = '#1E2124'   # Charcoal - body text on light, per the guide exactly
$COLOR_GRAY     = 'rgba(30,33,36,0.62)'   # Charcoal at reduced opacity - muted/secondary text
$COLOR_GREEN    = '#16A34A'   # semantic PASS / positive (not a brand color - status semantics)
$COLOR_RED      = '#DC2626'   # semantic FAIL / negative
$COLOR_AMBER    = '#D97706'   # semantic WARN
$COLOR_BLUE     = '#D4A94F'   # Sovereign Gold
$COLOR_BRAND    = '#071A33'   # Admiralty (decorative watermark)
$COLOR_GOLD     = '#D4A94F'   # Sovereign Gold
$COLOR_GOLD_DK  = '#B8902E'   # Sovereign Gold, darkened for text-sized contrast use
$COLOR_SIGNAL   = '#16C1C8'   # Signal Teal - sparing accent only (underlines, ticks)

# ---------------------------------------------------------------------------
# Load JSON
# ---------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $JsonPath)) {
    throw "JSON file not found: $JsonPath"
}
Write-Host "Reading assessment data from: $JsonPath" -ForegroundColor Cyan
$data = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json

# ---------------------------------------------------------------------------
# Resolve assessment tier (drives which report copy renders below - see
# "Tier-specific copy" further down). Defaults to Essential (the free,
# self-serve tier) for any JSON generated before AssessmentTier existed,
# matching what the live automated pipeline has always produced.
# ---------------------------------------------------------------------------
$AssessmentTier = ''
if ($data.PSObject.Properties.Name -contains 'AssessmentTier') { $AssessmentTier = [string]$data.AssessmentTier }
if ([string]::IsNullOrWhiteSpace($AssessmentTier)) { $AssessmentTier = 'Essential' }
$IsAdvancedTier = ($AssessmentTier -eq 'Advanced')

# ---------------------------------------------------------------------------
# Resolve output folder
# ---------------------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $OutputFolder = Split-Path -Parent (Resolve-Path -LiteralPath $JsonPath)
}
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

# ---------------------------------------------------------------------------
# Resolve customer name (auto-detect from default domain if blank)
# ---------------------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($CustomerName)) {
    $dom = [string]$data.DefaultDomain
    if (-not [string]::IsNullOrWhiteSpace($dom)) {
        $namePart = $dom -replace '\.onmicrosoft\.com$', ''
        $namePart = $namePart -replace '\..*$', ''
        if (-not [string]::IsNullOrWhiteSpace($namePart)) {
            $ti = (Get-Culture).TextInfo
            $CustomerName = $ti.ToTitleCase($namePart.ToLower())
        }
    }
    if ([string]::IsNullOrWhiteSpace($CustomerName)) { $CustomerName = 'Customer' }
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Enc([object]$s) {
    if ($null -eq $s) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$s)
}

# Report date
$reportDate = (Get-Date)
if ($data.GeneratedAt) {
    try { $reportDate = [datetime]::Parse([string]$data.GeneratedAt) } catch { $reportDate = Get-Date }
}
$reportDateStr = $reportDate.ToString('MMMM d, yyyy')

# ---------------------------------------------------------------------------
# SVG status icons (PASS / FAIL / MANUAL / WARN)
# ---------------------------------------------------------------------------
function Get-StatusIcon {
    param(
        [string]$Status,
        [int]$Size = 40
    )
    $s = ("$Status").Trim().ToUpper()
    switch -Regex ($s) {
        '^PASS$' {
            return "<svg width='$Size' height='$Size' viewBox='0 0 24 24'><circle cx='12' cy='12' r='11' fill='$COLOR_GREEN'/><polyline points='6.5,12.5 10.3,16 17,8' fill='none' stroke='#ffffff' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'/></svg>"
        }
        '^FAIL$' {
            return "<svg width='$Size' height='$Size' viewBox='0 0 24 24'><circle cx='12' cy='12' r='11' fill='none' stroke='$COLOR_RED' stroke-width='2'/><line x1='8' y1='8' x2='16' y2='16' stroke='$COLOR_RED' stroke-width='2' stroke-linecap='round'/><line x1='16' y1='8' x2='8' y2='16' stroke='$COLOR_RED' stroke-width='2' stroke-linecap='round'/></svg>"
        }
        '^WARN' {
            return "<svg width='$Size' height='$Size' viewBox='0 0 24 24'><circle cx='12' cy='12' r='11' fill='none' stroke='$COLOR_AMBER' stroke-width='2'/><line x1='7' y1='12' x2='17' y2='12' stroke='$COLOR_AMBER' stroke-width='2' stroke-linecap='round'/></svg>"
        }
        '^(MANUAL|NOTSET|NOT SET|NA|N/A|INFO)$' {
            return "<svg width='$Size' height='$Size' viewBox='0 0 24 24'><circle cx='12' cy='12' r='11' fill='none' stroke='$COLOR_GRAY' stroke-width='2'/><line x1='7' y1='12' x2='17' y2='12' stroke='$COLOR_GRAY' stroke-width='2' stroke-linecap='round'/></svg>"
        }
        default {
            return "<svg width='$Size' height='$Size' viewBox='0 0 24 24'><circle cx='12' cy='12' r='11' fill='none' stroke='$COLOR_GRAY' stroke-width='2'/></svg>"
        }
    }
}

# ---------------------------------------------------------------------------
# Small stat indicator icons (next to metric numbers)
#   ok   -> green check circle
#   warn -> orange triangle with !
#   bad  -> red triangle with !
#   info -> orange circle with !
# ---------------------------------------------------------------------------
function Get-StatIndicator {
    param([string]$Level, [int]$Size = 26)
    switch ($Level) {
        'ok' {
            return "<svg width='$Size' height='$Size' viewBox='0 0 24 24'><circle cx='12' cy='12' r='9.5' fill='none' stroke='$COLOR_GREEN' stroke-width='1.6'/><polyline points='7.5,12.5 10.8,15.7 16.5,9' fill='none' stroke='$COLOR_GREEN' stroke-width='1.6' stroke-linecap='round' stroke-linejoin='round'/></svg>"
        }
        'info' {
            return "<svg width='$Size' height='$Size' viewBox='0 0 24 24'><circle cx='12' cy='12' r='9.5' fill='none' stroke='$COLOR_AMBER' stroke-width='1.6'/><line x1='12' y1='7.5' x2='12' y2='13' stroke='$COLOR_AMBER' stroke-width='1.7' stroke-linecap='round'/><circle cx='12' cy='16.3' r='1' fill='$COLOR_AMBER'/></svg>"
        }
        'warn' {
            return "<svg width='$Size' height='$Size' viewBox='0 0 24 24'><path d='M12 3 L22 20 L2 20 Z' fill='none' stroke='$COLOR_AMBER' stroke-width='1.7' stroke-linejoin='round'/><line x1='12' y1='9' x2='12' y2='15' stroke='$COLOR_AMBER' stroke-width='1.7' stroke-linecap='round'/><circle cx='12' cy='17.6' r='1' fill='$COLOR_AMBER'/></svg>"
        }
        'bad' {
            return "<svg width='$Size' height='$Size' viewBox='0 0 24 24'><path d='M12 3 L22 20 L2 20 Z' fill='none' stroke='$COLOR_RED' stroke-width='1.7' stroke-linejoin='round'/><line x1='12' y1='9' x2='12' y2='15' stroke='$COLOR_RED' stroke-width='1.7' stroke-linecap='round'/><circle cx='12' cy='17.6' r='1' fill='$COLOR_RED'/></svg>"
        }
        default {
            return "<svg width='$Size' height='$Size' viewBox='0 0 24 24'><circle cx='12' cy='12' r='9.5' fill='none' stroke='$COLOR_GRAY' stroke-width='1.6'/></svg>"
        }
    }
}

# ---------------------------------------------------------------------------
# Severity badges
# ---------------------------------------------------------------------------
function Get-SeverityBadge {
    param([string]$Severity)
    $sev = ("$Severity").Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($sev)) { return '' }
    switch ($sev) {
        'critical' { return "<span class='badge' style='background:#dc2626;color:#fff;'>critical</span>" }
        'high'     { return "<span class='badge' style='border:1px solid #f97316;color:#f97316;'>high</span>" }
        'medium'   { return "<span class='badge' style='border:1px solid $COLOR_BLUE;color:$COLOR_BLUE;'>medium</span>" }
        'low'      { return "<span class='badge' style='border:1px solid rgba(30,33,36,0.55);color:rgba(30,33,36,0.55);'>low</span>" }
        default    { return "<span class='badge' style='border:1px solid rgba(30,33,36,0.55);color:rgba(30,33,36,0.55);'>$(Enc $sev)</span>" }
    }
}

# ---------------------------------------------------------------------------
# User / Device health metric icons (70px, stroke based)
# ---------------------------------------------------------------------------
function Get-MetricIcon {
    param([string]$Kind)
    # Iconography guideline: single-weight 1.5px stroke in Admiralty navy,
    # no fill - not the muted gray this used before official brand
    # guidelines were available to check against.
    $st = "fill='none' stroke='$COLOR_TEXT' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'"
    switch ($Kind) {
        'person'   { return "<svg width='40' height='40' viewBox='0 0 24 24'><circle cx='12' cy='8' r='4' $st/><path d='M4 21c0-4.4 3.6-7 8-7s8 2.6 8 7' $st/></svg>" }
        'admin'    { return "<svg width='40' height='40' viewBox='0 0 24 24'><path d='M6 10 L6 6 L9 8 L12 4 L15 8 L18 6 L18 10 Z' $st/><path d='M6 12 h12' $st/><path d='M5 20c0-3.3 3.1-5 7-5s7 1.7 7 5' $st/></svg>" }
        'lock'     { return "<svg width='40' height='40' viewBox='0 0 24 24'><circle cx='9' cy='8' r='3.5' $st/><path d='M2.5 20c0-3.3 2.9-5 6.5-5 1 0 2 .1 2.8.4' $st/><rect x='13' y='13' width='9' height='7' rx='1.2' $st/><path d='M15 13 v-2 a2.5 2.5 0 0 1 5 0 v2' $st/></svg>" }
        'lock-open'{ return "<svg width='40' height='40' viewBox='0 0 24 24'><rect x='5' y='11' width='14' height='10' rx='1.5' $st/><path d='M8 11 V7 a4 4 0 0 1 7.5-2' $st/><circle cx='12' cy='16' r='1.3' fill='$COLOR_TEXT' stroke='none'/></svg>" }
        'risky'    { return "<svg width='40' height='40' viewBox='0 0 24 24'><circle cx='10' cy='8' r='3.5' $st/><path d='M3 20c0-3.6 3.1-6 7-6' $st/><line x1='15' y1='14' x2='21' y2='20' $st/><line x1='21' y1='14' x2='15' y2='20' $st/></svg>" }
        'devices'  { return "<svg width='40' height='40' viewBox='0 0 24 24'><rect x='2' y='4' width='14' height='10' rx='1' $st/><path d='M6 18 h6' $st/><path d='M9 14 v4' $st/><rect x='16' y='9' width='6' height='11' rx='1.2' $st/></svg>" }
        'doc-alert'{ return "<svg width='40' height='40' viewBox='0 0 24 24'><path d='M7 3 h7 l5 5 v13 a1 1 0 0 1-1 1 H7 a1 1 0 0 1-1-1 V4 a1 1 0 0 1 1-1 Z' $st/><path d='M14 3 v5 h5' $st/><line x1='12.5' y1='11' x2='12.5' y2='15.5' $st/><circle cx='12.5' cy='18' r='.9' fill='$COLOR_TEXT' stroke='none'/></svg>" }
        'calendar' { return "<svg width='40' height='40' viewBox='0 0 24 24'><rect x='3' y='5' width='18' height='16' rx='1.5' $st/><path d='M3 9 h18' $st/><path d='M8 3 v4 M16 3 v4' $st/><line x1='12' y1='12' x2='12' y2='16' $st/><circle cx='12' cy='18.3' r='.9' fill='$COLOR_TEXT' stroke='none'/></svg>" }
        default    { return "<svg width='40' height='40' viewBox='0 0 24 24'><circle cx='12' cy='12' r='9' $st/></svg>" }
    }
}

# ---------------------------------------------------------------------------
# Sharing capability -> human readable
# ---------------------------------------------------------------------------
function Get-SharingLabel {
    param([string]$Value)
    switch (("$Value").Trim()) {
        'disabled'                        { return 'Disabled' }
        'existingExternalUserSharingOnly' { return 'Existing Guests Only' }
        'externalUserSharingOnly'         { return 'New and Existing Guests' }
        'externalUserAndGuestSharing'     { return 'Anyone' }
        default                           { return "$Value" }
    }
}
function Get-SharingDescription {
    param([string]$Value)
    switch (("$Value").Trim()) {
        'disabled'                        { return 'External sharing is turned off. Only people inside the organization can access content.' }
        'existingExternalUserSharingOnly' { return 'Sharing is limited to guests who already exist in the directory.' }
        'externalUserSharingOnly'         { return 'Content can be shared with new and existing authenticated guests.' }
        'externalUserAndGuestSharing'     { return 'By default, links are generated which can be accessed by anyone internal or external to the organization.' }
        default                           { return 'Review the SharePoint / OneDrive external sharing configuration.' }
    }
}
function Get-SharingSeverity {
    param([string]$Value)
    switch (("$Value").Trim()) {
        'disabled'                        { return 'PASS' }
        'existingExternalUserSharingOnly' { return 'PASS' }
        'externalUserSharingOnly'         { return 'WARN' }
        'externalUserAndGuestSharing'     { return 'FAIL' }
        default                           { return 'WARN' }
    }
}

# ---------------------------------------------------------------------------
# Static parent-title map (CheckId -> group heading)
# ---------------------------------------------------------------------------
$ParentTitles = @{
    '1.1'='Multi-factor authentication is enforced for all users'
    '1.2'='MFA is required for all Admins'
    '1.3'='Legacy Authentication is blocked'
    '1.4'='Break Glass users are created for emergency access'
    '1.5'='Ensure that between two and four global admins are designated'
    '1.6'='Highly privileged accounts shall be cloud-only'
    '1.7'='Non-admin users shall be prevented from providing consent to 3rd party applications'
    '1.8'='Guest users have limited access to properties and memberships of directory objects'
    '1.9'='Passwords shall not expire'
    '1.10'='MFA shall be required to enroll devices to Azure AD'
    '1.11'='Local Administrator settings are configured for device joins'
    '1.12'='Dormant Accounts are disabled with 45 days of Inactivity'
    '1.13'='Browser Sessions are limited for Privileged Users'
    '1.14'='Devices shall be deleted that have not checked in for over 30 days'
    '1.15'='All corporate approved applications are cataloged and periodically reviewed'
    '1.16'='Dynamic Groups are leveraged for automated group management'
    '1.17'='MFA Shall be required for Intune Enrollment'
    '1.18'='Require Managed Devices for Sign in'
    '1.19'='Device Compliance is required for access to resources'
    '1.20'='Require Phishing Resistant MFA for Admins'
    '1.21'='High risk users and sign-ins are blocked'
    '1.22'='Privileged Identity Management (PIM) is configured for JIT access'
    '1.23'='Microsoft Sentinel is configured to ingest logs from Entra and Defender'
    '1.24'='Ensure the device code sign-in flow is blocked'
    '2.1'='Email authentication is configured for all domains (SPF, DKIM, DMARC)'
    '2.2'='Anti-phishing policies are configured'
    '2.3'='Anti-spam policies are configured'
    '2.4'='External email tagging is enabled'
    '2.5'='Auto-forwarding of email to external domains is blocked'
    '2.6'='Audit logging is enabled for all mailboxes'
    '3.1'='External User Access SHALL Be Restricted'
    '3.2'='External Participants SHOULD NOT Be Enabled to Request Control of Shared Desktops or Windows in Meetings'
    '3.3'='Anonymous Users SHALL NOT Be Enabled to Start Meetings'
    '3.4'='Automatic Admittance to Meetings SHOULD Be Restricted'
    '3.5'='Unmanaged users SHALL NOT be enabled to initiate contact with internal users'
    '3.6'='Contact with Skype Users SHALL Be Blocked'
    '3.7'='File Sharing and File Storage Options shall be blocked'
    '4.1'='Automated patching is performed on all devices'
    '4.2'='Managed devices are enrolled in MDM'
    '4.3'='Personal Devices should be restricted from enrolling into the MDM solution'
    '4.4'='Security Baselines should be configured for Windows Devices'
    '4.5'='Devices compliance policies shall be configured for every supported device platform'
    '4.6'='All devices have drive encryption applied'
    '4.7'='Lockout screen and password settings shall be configured for each device'
    '4.8'='App Protection policies should be created for mobile devices'
    '4.9'='Approved 3rd party applications are deployed and patched'
    '4.10'='Local Administrators passwords are managed with LAPS'
    '5.1'='Default sharing settings are set for New and Existing Guest'
    '5.2'='Expiration Dates are set for Anyone links'
    '6.1'='Security Awareness training is conducted at least once per year'
    '6.2'='Anti-virus protections are applied to all devices'
    '6.3'='Endpoint detection and response software is running on all devices'
    '6.4'='Firewall protections configured on devices'
    '6.5'='Safe Links policies are configured'
    '6.6'='Safe Attachment policies are configured'
    '6.7'='Tamper Protection is configured'
    '6.8'='Attack Surface reduction rules are configured'
    '6.9'='Defender for Cloud Apps is configured to monitor applications on the network'
    '7.1'='Periodic backups are performed for email, files, and Servers'
    '7.2'='Audit Logging SHALL Be Enabled'
    '7.3'='Retention Policies are configured'
    '7.4'='Sensitivity Labels are configured'
    '7.5'='Data Loss Prevention Policies are configured'
}

# Section display titles + order
$SectionOrder = @('1-EntraID','2-Exchange','3-Teams','4-Intune','5-SharePoint','6-Defender','7-Purview')
$SectionTitles = @{
    '1-EntraID'    = '1 - Entra ID'
    '2-Exchange'   = '2 - Exchange'
    '3-Teams'      = '3 - Teams'
    '4-Intune'     = '4 - Intune'
    '5-SharePoint' = '5 - SharePoint and OneDrive'
    '6-Defender'   = '6 - Defender'
    '7-Purview'    = '7 - Purview'
}

# ---------------------------------------------------------------------------
# Page watermark - the real Camelot crest mark (website/public/brand-assets/
# mark/camelot-mark-transparent.svg), embedded verbatim and placed on every
# page at low opacity via CSS (.page-watermark). Replaces the earlier
# decorative semicircle, which didn't carry any actual brand identity.
# ---------------------------------------------------------------------------
$Watermark = @'
<div class='page-watermark'>
<svg xmlns='http://www.w3.org/2000/svg' viewBox='251 40 698 698'>
<defs>
<linearGradient id='cdgGold' x1='0' y1='0' x2='1' y2='1'>
<stop offset='0%' stop-color='#B98B25'/><stop offset='18%' stop-color='#F2DC93'/>
<stop offset='38%' stop-color='#C9992C'/><stop offset='58%' stop-color='#E9C864'/>
<stop offset='80%' stop-color='#B0801C'/><stop offset='100%' stop-color='#E4C169'/>
</linearGradient>
<linearGradient id='cdgSilver' x1='0' y1='0' x2='1' y2='1'>
<stop offset='0%' stop-color='#9AA3AF'/><stop offset='22%' stop-color='#F2F5F9'/>
<stop offset='50%' stop-color='#AEB7C4'/><stop offset='76%' stop-color='#DFE5ED'/>
<stop offset='100%' stop-color='#CBD3DE'/>
</linearGradient>
</defs>
<g fill='url(#cdgGold)' transform='translate(290,40)'><path fill-rule='evenodd' d='M310,0C413.624,69.636 430.121,69.639 620,99.2L620,258.1C620,540.55 482.74,595.4 310,698C137.26,595.4 0,540.55 0,258.1L0,99.2C189.879,69.639 206.376,69.636 310,0ZM310,49.56C230.181,102.466 185.914,113.048 41.33,134.61L41.33,266.63C41.33,502.084 153.968,560.428 310,649.97C466.032,560.428 578.67,502.084 578.67,266.63L578.67,134.61C434.086,113.048 389.819,102.466 310,49.56Z'/></g>
<g fill='url(#cdgSilver)'><path transform='translate(436.68,511.12)' d='M302.16 -122.28 306.29 -50.18Q296.93 -34.71 281.40 -22.08Q265.88 -9.45 242.96 -1.79Q220.05 5.86 188.61 5.86Q137.43 5.44 99.06 -12.77Q60.69 -30.97 39.73 -64.98Q18.77 -98.98 18.77 -146.99Q18.77 -193.91 39.39 -228.04Q60.02 -262.17 98.34 -280.59Q136.67 -299.01 189.85 -299.01Q223.39 -299.01 249.83 -292.24Q276.27 -285.47 294.19 -276.12L294.94 -204.36H291.17Q279.92 -245.79 254.24 -263.57Q228.56 -281.36 194.18 -281.36Q160.88 -281.36 136.92 -265.12Q112.96 -248.89 100.00 -219.09Q87.04 -189.29 87.04 -147.92Q87.04 -106.55 99.50 -76.28Q111.95 -46.02 134.73 -29.32Q157.50 -12.62 188.61 -11.79Q220.89 -11.79 243.05 -22.25Q265.20 -32.71 278.81 -57.03Q292.42 -81.35 298.81 -122.28Z'/></g>
</svg>
</div>
'@

# ---------------------------------------------------------------------------
# Gauge (270 deg arc)
# ---------------------------------------------------------------------------
function Get-GaugeSvg {
    param([int]$Percentage, [double]$Current, [double]$Max)
    $r = 85
    $cx = 110; $cy = 100
    $circ = 2 * [math]::PI * $r          # ~534.07
    $trackLen = $circ * 0.75             # 270 deg  ~400.55
    $gap = $circ - $trackLen             # ~134.52
    $p = [math]::Max(0, [math]::Min(100, $Percentage))
    $fillLen = [math]::Round($trackLen * ($p / 100.0), 2)
    $trackLen = [math]::Round($trackLen, 2)
    $gap = [math]::Round($gap, 2)
    $rest = [math]::Round($circ - $fillLen, 2)
    $curStr = ('{0:0.#}' -f $Current)
    $maxStr = ('{0:0.#}' -f $Max)
    return @"
<svg class='gauge' width='260' height='210' viewBox='0 0 220 180'>
  <g transform='rotate(135 $cx $cy)'>
    <circle cx='$cx' cy='$cy' r='$r' fill='none' stroke='$COLOR_BORDER' stroke-width='14'
            stroke-dasharray='$trackLen $gap' stroke-linecap='round'/>
    <circle cx='$cx' cy='$cy' r='$r' fill='none' stroke='$COLOR_GREEN' stroke-width='14'
            stroke-dasharray='$fillLen $rest' stroke-linecap='round'/>
  </g>
  <text x='$cx' y='$($cy+2)' text-anchor='middle' fill='$COLOR_TEXT' font-size='38' font-weight='700'>$p%</text>
  <text x='$cx' y='$($cy+26)' text-anchor='middle' fill='$COLOR_GRAY' font-size='13'>$curStr / $maxStr</text>
</svg>
"@
}

# ---------------------------------------------------------------------------
# Line chart (score history)
# ---------------------------------------------------------------------------
function Get-LineChartSvg {
    param($History)
    $pts = @()
    if ($History) {
        foreach ($h in $History) {
            $d = $null
            try { $d = [datetime]::Parse([string]$h.createdDateTime) } catch { continue }
            $cur = 0.0; $mx = 1.0
            try { $cur = [double]$h.currentScore } catch {}
            try { $mx = [double]$h.maxScore } catch {}
            if ($mx -le 0) { $mx = 1 }
            $pct = [math]::Round(($cur / $mx) * 100.0, 1)
            $pts += [pscustomobject]@{ Date = $d; Pct = $pct }
        }
    }
    $pts = $pts | Sort-Object Date
    if ($pts.Count -gt 30) { $pts = $pts | Select-Object -Last 30 }

    $W = 460; $H = 190
    $padL = 45; $padR = 15; $padT = 15; $padB = 34
    $plotW = $W - $padL - $padR
    $plotH = $H - $padT - $padB

    if ($pts.Count -eq 0) {
        return "<svg width='$W' height='$H' viewBox='0 0 $W $H'><text x='$($W/2)' y='$($H/2)' text-anchor='middle' fill='$COLOR_GRAY' font-size='13'>No score history available</text></svg>"
    }

    $vals = $pts | ForEach-Object { $_.Pct }
    $minV = ($vals | Measure-Object -Minimum).Minimum
    $maxV = ($vals | Measure-Object -Maximum).Maximum
    if ($maxV -eq $minV) { $minV -= 5; $maxV += 5 }
    $range = $maxV - $minV
    $minV = [math]::Floor($minV - $range * 0.10)
    $maxV = [math]::Ceiling($maxV + $range * 0.10)
    if ($minV -lt 0) { $minV = 0 }
    $range = $maxV - $minV
    if ($range -le 0) { $range = 1 }

    $n = $pts.Count
    function _x([int]$i) { if ($n -le 1) { return $padL + $plotW/2 } return $padL + ($plotW * $i / ($n - 1)) }
    function _y([double]$v) { return $padT + $plotH - (($v - $minV) / $range * $plotH) }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg width='$W' height='$H' viewBox='0 0 $W $H'>")

    # grid lines + y labels (min, mid, max)
    $labels = @($minV, [math]::Round(($minV+$maxV)/2), $maxV)
    foreach ($lv in ($labels | Sort-Object -Unique)) {
        $gy = [math]::Round((_y $lv), 1)
        [void]$sb.Append("<line x1='$padL' y1='$gy' x2='$($W-$padR)' y2='$gy' stroke='$COLOR_BORDER' stroke-width='1'/>")
        [void]$sb.Append("<text x='$($padL-8)' y='$($gy+4)' text-anchor='end' fill='$COLOR_GRAY' font-size='11'>$lv</text>")
    }

    # polyline
    $poly = ($pts | ForEach-Object -Begin { $i = 0 } -Process {
        $x = [math]::Round((_x $i), 1); $y = [math]::Round((_y $_.Pct), 1); $i++
        "$x,$y"
    }) -join ' '
    [void]$sb.Append("<polyline points='$poly' fill='none' stroke='$COLOR_GREEN' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'/>")

    # x-axis first / last labels
    $firstLbl = $pts[0].Date.ToString('MMM d')
    $lastLbl  = $pts[$pts.Count-1].Date.ToString('MMM d')
    [void]$sb.Append("<text x='$padL' y='$($H-10)' text-anchor='start' fill='$COLOR_GRAY' font-size='11'>$firstLbl</text>")
    [void]$sb.Append("<text x='$($W-$padR)' y='$($H-10)' text-anchor='end' fill='$COLOR_GRAY' font-size='11'>$lastLbl</text>")

    [void]$sb.Append('</svg>')
    return $sb.ToString()
}

# ---------------------------------------------------------------------------
# Logo (embed or text placeholder)
# ---------------------------------------------------------------------------
function Get-LogoHtml {
    if (-not [string]::IsNullOrWhiteSpace($LogoPath) -and (Test-Path -LiteralPath $LogoPath)) {
        $ext = ([System.IO.Path]::GetExtension($LogoPath)).ToLower()
        if ($ext -eq '.svg') {
            $svg = Get-Content -LiteralPath $LogoPath -Raw
            return "<div class='logo'>$svg</div>"
        } else {
            $bytes = [System.IO.File]::ReadAllBytes($LogoPath)
            $b64 = [Convert]::ToBase64String($bytes)
            $mime = switch ($ext) { '.png' {'image/png'} '.jpg' {'image/jpeg'} '.jpeg' {'image/jpeg'} '.gif' {'image/gif'} default {'image/png'} }
            return "<div class='logo'><img src='data:$mime;base64,$b64' alt='logo'/></div>"
        }
    }
    # Real Camelot Digital Group lockup (mark + wordmark), the same asset
    # the live site itself uses (website/public/brand-assets/lockup/
    # camelot-lockup-on-light.svg) - embedded verbatim rather than the
    # earlier hand-drawn "CDG shield" approximation, since this script
    # has no filesystem access to the website's own assets at runtime
    # (it runs inside a separate Azure Automation sandbox).
    $lockup = @'
<svg xmlns="http://www.w3.org/2000/svg" width="473" height="140" viewBox="0 0 473.25 140.00" role="img" aria-label="Camelot Digital Group">
<title>Camelot Digital Group</title>
<defs>
<linearGradient id="cdgGold" x1="0" y1="0" x2="1" y2="1">
<stop offset="0%" stop-color="#B98B25"/><stop offset="18%" stop-color="#F2DC93"/>
<stop offset="38%" stop-color="#C9992C"/><stop offset="58%" stop-color="#E9C864"/>
<stop offset="80%" stop-color="#B0801C"/><stop offset="100%" stop-color="#E4C169"/>
</linearGradient>
<linearGradient id="cdgSilver" x1="0" y1="0" x2="1" y2="1">
<stop offset="0%" stop-color="#9AA3AF"/><stop offset="22%" stop-color="#F2F5F9"/>
<stop offset="50%" stop-color="#AEB7C4"/><stop offset="76%" stop-color="#DFE5ED"/>
<stop offset="100%" stop-color="#CBD3DE"/>
</linearGradient>
</defs>
<g transform="translate(-50.842,-3.564) scale(0.18911)"><g fill="url(#cdgGold)"><g transform="translate(290,40)">
<path fill-rule="evenodd" d="M310,0C413.624,69.636 430.121,69.639 620,99.2L620,258.1C620,540.55 482.74,595.4 310,698C137.26,595.4 0,540.55 0,258.1L0,99.2C189.879,69.639 206.376,69.636 310,0ZM310,49.56C230.181,102.466 185.914,113.048 41.33,134.61L41.33,266.63C41.33,502.084 153.968,560.428 310,649.97C466.032,560.428 578.67,502.084 578.67,266.63L578.67,134.61C434.086,113.048 389.819,102.466 310,49.56Z"/>
</g></g><g fill="#071A33"><path transform="translate(436.68,511.12)" d="M302.16 -122.28 306.29 -50.18Q296.93 -34.71 281.40 -22.08Q265.88 -9.45 242.96 -1.79Q220.05 5.86 188.61 5.86Q137.43 5.44 99.06 -12.77Q60.69 -30.97 39.73 -64.98Q18.77 -98.98 18.77 -146.99Q18.77 -193.91 39.39 -228.04Q60.02 -262.17 98.34 -280.59Q136.67 -299.01 189.85 -299.01Q223.39 -299.01 249.83 -292.24Q276.27 -285.47 294.19 -276.12L294.94 -204.36H291.17Q279.92 -245.79 254.24 -263.57Q228.56 -281.36 194.18 -281.36Q160.88 -281.36 136.92 -265.12Q112.96 -248.89 100.00 -219.09Q87.04 -189.29 87.04 -147.92Q87.04 -106.55 99.50 -76.28Q111.95 -46.02 134.73 -29.32Q157.50 -12.62 188.61 -11.79Q220.89 -11.79 243.05 -22.25Q265.20 -32.71 278.81 -57.03Q292.42 -81.35 298.81 -122.28Z"/></g></g>
<g fill="#071A33"><path transform="translate(151.25,80.00)" d="M47.42 -19.19 48.06 -7.87Q46.59 -5.45 44.16 -3.47Q41.72 -1.48 38.13 -0.28Q34.53 0.92 29.60 0.92Q21.57 0.85 15.54 -2.00Q9.52 -4.86 6.23 -10.20Q2.95 -15.53 2.95 -23.07Q2.95 -30.43 6.18 -35.78Q9.42 -41.14 15.43 -44.03Q21.45 -46.92 29.79 -46.92Q35.05 -46.92 39.20 -45.86Q43.35 -44.80 46.16 -43.33L46.28 -32.07H45.69Q43.93 -38.57 39.90 -41.36Q35.87 -44.15 30.47 -44.15Q25.25 -44.15 21.49 -41.60Q17.73 -39.06 15.69 -34.38Q13.66 -29.70 13.66 -23.21Q13.66 -16.72 15.61 -11.97Q17.57 -7.22 21.14 -4.60Q24.72 -1.98 29.60 -1.85Q34.66 -1.85 38.14 -3.49Q41.62 -5.13 43.75 -8.95Q45.89 -12.76 46.89 -19.19ZM70.08 -47.32 91.53 -0.39H79.76L66.10 -34.65ZM55.96 -4.80Q55.42 -3.48 55.74 -2.56Q56.07 -1.64 56.87 -1.15Q57.66 -0.66 58.45 -0.66H59.04V0.00H44.96V-0.66Q44.96 -0.66 45.25 -0.66Q45.55 -0.66 45.55 -0.66Q47.06 -0.66 48.67 -1.61Q50.28 -2.56 51.33 -4.80ZM70.08 -47.32 70.32 -39.29 54.02 -0.20H49.24L66.26 -38.72Q66.46 -39.07 66.94 -40.17Q67.43 -41.28 68.00 -42.66Q68.57 -44.04 69.01 -45.32Q69.45 -46.59 69.49 -47.32ZM77.00 -15.67V-12.90H57.62V-15.67ZM77.96 -4.80H89.49Q90.56 -2.56 92.16 -1.61Q93.76 -0.66 95.27 -0.66Q95.27 -0.66 95.54 -0.66Q95.80 -0.66 95.80 -0.66V0.00H74.87V-0.66H75.46Q76.75 -0.66 77.76 -1.74Q78.76 -2.83 77.96 -4.80ZM140.73 -46.92 141.26 -39.76 121.73 -5.94Q121.73 -5.94 121.10 -4.79Q120.47 -3.63 119.84 -1.96Q119.21 -0.30 119.18 1.32H118.56L115.96 -4.42ZM96.17 -4.80V0.00H88.87V-0.66Q88.94 -0.66 89.43 -0.66Q89.93 -0.66 89.93 -0.66Q91.70 -0.66 93.08 -1.74Q94.46 -2.83 94.72 -4.80ZM99.07 -3.75Q99.07 -3.68 99.07 -3.61Q99.07 -3.55 99.07 -3.42Q99.07 -2.37 99.83 -1.48Q100.58 -0.59 101.63 -0.59H102.66V0.00H98.61V-3.75ZM100.17 -46.92H100.79L103.10 -40.33L98.66 0.00H94.12ZM100.79 -46.92 122.27 -11.65 118.56 1.32 99.10 -31.34ZM141.32 -46.92 147.57 0.00H137.02L133.69 -30.21L140.73 -46.92ZM145.52 -4.80H146.94Q147.27 -2.83 148.63 -1.74Q149.99 -0.66 151.74 -0.66Q151.74 -0.66 152.24 -0.66Q152.75 -0.66 152.79 -0.66V0.00H145.52ZM136.61 -3.75H137.07V0.00H133.02V-0.59H134.05Q135.13 -0.59 135.87 -1.48Q136.61 -2.37 136.61 -3.42Q136.61 -3.55 136.61 -3.61Q136.61 -3.68 136.61 -3.75ZM165.33 -46.00V0.00H155.47V-46.00ZM182.67 -2.84 184.27 0.00H165.13V-2.84ZM180.24 -24.06V-21.29H165.13V-24.06ZM183.47 -46.00V-43.16H165.13V-46.00ZM187.48 -13.44 184.46 0.00H171.56L173.93 -2.84Q177.21 -2.84 179.73 -4.12Q182.25 -5.40 184.03 -7.79Q185.81 -10.18 186.84 -13.44ZM180.24 -21.42V-15.09H179.59V-15.67Q179.59 -18.06 178.01 -19.66Q176.43 -21.25 174.04 -21.29V-21.42ZM180.24 -30.26V-23.92H174.04V-24.06Q176.43 -24.12 178.03 -25.73Q179.63 -27.34 179.59 -29.74V-30.26ZM183.47 -43.36V-35.71H182.81V-36.50Q182.81 -39.51 181.02 -41.34Q179.22 -43.16 176.14 -43.23V-43.36ZM183.47 -47.18V-45.28L175.45 -46.00Q176.90 -46.00 178.51 -46.20Q180.12 -46.39 181.50 -46.66Q182.88 -46.92 183.47 -47.18ZM155.67 -4.80V0.00H150.41V-0.66Q150.41 -0.66 150.84 -0.66Q151.26 -0.66 151.26 -0.66Q152.97 -0.66 154.19 -1.87Q155.40 -3.09 155.47 -4.80ZM155.67 -41.20H155.47Q155.43 -42.91 154.19 -44.13Q152.96 -45.34 151.26 -45.34Q151.26 -45.34 150.85 -45.34Q150.44 -45.34 150.44 -45.34L150.41 -46.00H155.67ZM202.41 -46.00V0.00H192.56V-46.00ZM219.76 -2.84 221.35 0.00H202.22V-2.84ZM224.56 -13.44 221.55 0.00H208.65L211.01 -2.84Q214.30 -2.84 216.82 -4.12Q219.34 -5.40 221.12 -7.79Q222.90 -10.18 223.92 -13.44ZM192.75 -4.80V0.00H187.50V-0.66Q187.50 -0.66 187.92 -0.66Q188.35 -0.66 188.35 -0.66Q190.06 -0.66 191.28 -1.87Q192.49 -3.09 192.56 -4.80ZM192.75 -41.20H192.56Q192.49 -42.91 191.28 -44.13Q190.06 -45.34 188.35 -45.34Q188.35 -45.34 187.92 -45.34Q187.50 -45.34 187.50 -45.34V-46.00H192.75ZM202.22 -41.20V-46.00H207.47V-45.34Q207.41 -45.34 207.01 -45.34Q206.62 -45.34 206.62 -45.34Q204.91 -45.34 203.69 -44.13Q202.48 -42.91 202.41 -41.20ZM249.88 -46.92Q257.79 -46.92 263.61 -44.01Q269.43 -41.10 272.60 -35.75Q275.77 -30.40 275.77 -23.00Q275.77 -15.64 272.60 -10.27Q269.43 -4.90 263.61 -1.99Q257.79 0.92 249.88 0.92Q242.00 0.92 236.20 -1.99Q230.39 -4.90 227.22 -10.25Q224.05 -15.60 224.05 -23.00Q224.05 -30.36 227.22 -35.73Q230.39 -41.10 236.20 -44.01Q242.00 -46.92 249.88 -46.92ZM249.88 -1.85Q254.55 -1.85 257.92 -4.43Q261.30 -7.01 263.13 -11.74Q264.97 -16.47 264.97 -23.00Q264.97 -29.53 263.13 -34.26Q261.30 -38.99 257.92 -41.57Q254.55 -44.15 249.88 -44.15Q245.28 -44.15 241.90 -41.57Q238.53 -38.99 236.69 -34.26Q234.86 -29.53 234.86 -23.00Q234.86 -16.47 236.69 -11.74Q238.53 -7.01 241.90 -4.43Q245.28 -1.85 249.88 -1.85ZM301.03 -45.80V0.00H291.17V-45.80ZM317.12 -46.07V-43.30H275.08V-46.07ZM317.12 -43.49V-35.78L316.46 -35.84V-36.56Q316.46 -39.60 314.66 -41.42Q312.85 -43.23 309.80 -43.30V-43.49ZM317.12 -47.25V-45.34L309.10 -46.07Q310.55 -46.07 312.16 -46.26Q313.77 -46.46 315.15 -46.72Q316.53 -46.99 317.12 -47.25ZM291.37 -4.80V0.00H286.11V-0.66Q286.11 -0.66 286.54 -0.66Q286.97 -0.66 286.97 -0.66Q288.68 -0.66 289.89 -1.87Q291.11 -3.09 291.17 -4.80ZM300.83 -4.80H301.03Q301.10 -3.09 302.31 -1.87Q303.53 -0.66 305.24 -0.66Q305.24 -0.66 305.66 -0.66Q306.09 -0.66 306.09 -0.66V0.00H300.83ZM282.41 -43.49V-43.30Q279.33 -43.23 277.54 -41.42Q275.74 -39.60 275.74 -36.56V-35.84L275.08 -35.78V-43.49ZM275.08 -47.25Q275.74 -46.99 277.09 -46.72Q278.44 -46.46 280.08 -46.26Q281.72 -46.07 283.10 -46.07L275.08 -45.34Z"/></g>
<g fill="#16C1C8"><path transform="translate(151.25,116.00)" d="M9.84 -17.00Q12.70 -17.00 14.79 -15.97Q16.88 -14.95 18.01 -13.03Q19.15 -11.12 19.15 -8.50Q19.15 -5.88 18.01 -3.97Q16.88 -2.06 14.79 -1.03Q12.70 0.00 9.84 0.00H4.05L4.03 -1.02Q4.73 -1.02 5.50 -1.02Q6.26 -1.02 6.99 -1.02Q7.71 -1.02 8.29 -1.02Q8.87 -1.02 9.20 -1.02Q9.54 -1.02 9.54 -1.02Q11.29 -1.02 12.55 -1.94Q13.82 -2.85 14.50 -4.52Q15.19 -6.20 15.19 -8.50Q15.19 -10.80 14.50 -12.48Q13.82 -14.15 12.55 -15.06Q11.28 -15.98 9.54 -15.98Q9.54 -15.98 9.19 -15.98Q8.83 -15.98 8.23 -15.98Q7.62 -15.98 6.87 -15.98Q6.12 -15.98 5.32 -15.98Q4.52 -15.98 3.78 -15.98V-17.00ZM6.53 -17.00V0.00H2.88V-17.00ZM2.96 -1.77V0.00H1.01V-0.24Q1.01 -0.24 1.17 -0.24Q1.33 -0.24 1.33 -0.24Q1.96 -0.24 2.41 -0.69Q2.86 -1.14 2.88 -1.77ZM2.96 -15.23H2.88Q2.88 -15.86 2.42 -16.31Q1.96 -16.76 1.33 -16.76Q1.33 -16.76 1.18 -16.76Q1.04 -16.76 1.04 -16.76L1.01 -17.00H2.96ZM35.67 -17.00V0.00H32.03V-17.00ZM32.10 -1.77V0.00H30.16V-0.24Q30.16 -0.24 30.32 -0.24Q30.48 -0.24 30.48 -0.24Q31.11 -0.24 31.56 -0.69Q32.01 -1.14 32.03 -1.77ZM32.10 -15.23H32.03Q32.01 -15.86 31.56 -16.31Q31.11 -16.76 30.48 -16.76Q30.48 -16.76 30.32 -16.76Q30.16 -16.76 30.16 -16.76V-17.00H32.10ZM35.60 -1.77H35.67Q35.70 -1.14 36.15 -0.69Q36.60 -0.24 37.23 -0.24Q37.23 -0.24 37.37 -0.24Q37.52 -0.24 37.54 -0.24V0.00H35.60ZM35.60 -15.23V-17.00H37.54V-16.76Q37.52 -16.76 37.37 -16.76Q37.23 -16.76 37.23 -16.76Q36.60 -16.76 36.15 -16.31Q35.70 -15.86 35.67 -15.23ZM65.91 -6.80V-2.79Q65.34 -2.06 64.18 -1.34Q63.03 -0.61 61.42 -0.13Q59.81 0.34 57.85 0.34Q55.10 0.33 53.03 -0.75Q50.97 -1.82 49.82 -3.81Q48.67 -5.79 48.67 -8.53Q48.67 -11.24 49.81 -13.21Q50.94 -15.19 53.04 -16.26Q55.14 -17.34 58.00 -17.34Q59.33 -17.34 60.60 -17.16Q61.87 -16.98 62.94 -16.68Q64.00 -16.39 64.71 -16.01L64.75 -11.85H64.53Q64.08 -13.51 63.21 -14.48Q62.33 -15.46 61.18 -15.89Q60.03 -16.32 58.73 -16.32Q56.75 -16.32 55.39 -15.36Q54.03 -14.39 53.33 -12.62Q52.63 -10.85 52.63 -8.38Q52.63 -6.04 53.28 -4.31Q53.92 -2.58 55.16 -1.63Q56.40 -0.69 58.17 -0.68Q59.08 -0.68 59.91 -0.96Q60.74 -1.24 61.37 -1.78Q62.00 -2.32 62.27 -3.09L62.29 -6.80Q62.29 -7.89 61.05 -7.89H60.66V-8.14H67.49V-7.89H67.12Q65.86 -7.89 65.91 -6.80ZM83.45 -17.00V0.00H79.81V-17.00ZM79.88 -1.77V0.00H77.94V-0.24Q77.94 -0.24 78.10 -0.24Q78.26 -0.24 78.26 -0.24Q78.89 -0.24 79.34 -0.69Q79.79 -1.14 79.81 -1.77ZM79.88 -15.23H79.81Q79.79 -15.86 79.34 -16.31Q78.89 -16.76 78.26 -16.76Q78.26 -16.76 78.10 -16.76Q77.94 -16.76 77.94 -16.76V-17.00H79.88ZM83.38 -1.77H83.45Q83.48 -1.14 83.93 -0.69Q84.38 -0.24 85.01 -0.24Q85.01 -0.24 85.15 -0.24Q85.30 -0.24 85.32 -0.24V0.00H83.38ZM83.38 -15.23V-17.00H85.32V-16.76Q85.30 -16.76 85.15 -16.76Q85.01 -16.76 85.01 -16.76Q84.38 -16.76 83.93 -16.31Q83.48 -15.86 83.45 -15.23ZM105.29 -16.93V0.00H101.65V-16.93ZM111.24 -17.02V-16.00H95.70V-17.02ZM111.24 -16.07V-13.22L110.99 -13.25V-13.51Q110.99 -14.63 110.33 -15.31Q109.66 -15.98 108.53 -16.00V-16.07ZM111.24 -17.46V-16.76L108.27 -17.02Q108.81 -17.02 109.40 -17.10Q110.00 -17.17 110.51 -17.27Q111.02 -17.36 111.24 -17.46ZM101.72 -1.77V0.00H99.78V-0.24Q99.78 -0.24 99.94 -0.24Q100.09 -0.24 100.09 -0.24Q100.72 -0.24 101.17 -0.69Q101.62 -1.14 101.65 -1.77ZM105.22 -1.77H105.29Q105.31 -1.14 105.76 -0.69Q106.21 -0.24 106.84 -0.24Q106.84 -0.24 107.00 -0.24Q107.16 -0.24 107.16 -0.24V0.00H105.22ZM98.41 -16.07V-16.00Q97.27 -15.98 96.61 -15.31Q95.94 -14.63 95.94 -13.51V-13.25L95.70 -13.22V-16.07ZM95.70 -17.46Q95.94 -17.36 96.44 -17.27Q96.94 -17.17 97.55 -17.10Q98.15 -17.02 98.66 -17.02L95.70 -16.76ZM129.20 -17.49 137.13 -0.15H132.78L127.73 -12.81ZM123.98 -1.77Q123.78 -1.29 123.90 -0.95Q124.02 -0.61 124.32 -0.43Q124.61 -0.24 124.90 -0.24H125.12V0.00H119.91V-0.24Q119.91 -0.24 120.02 -0.24Q120.13 -0.24 120.13 -0.24Q120.69 -0.24 121.29 -0.60Q121.88 -0.95 122.27 -1.77ZM129.20 -17.49 129.29 -14.52 123.26 -0.07H121.50L127.79 -14.31Q127.86 -14.44 128.04 -14.85Q128.22 -15.25 128.43 -15.76Q128.64 -16.28 128.80 -16.75Q128.97 -17.22 128.98 -17.49ZM131.75 -5.79V-4.77H124.60V-5.79ZM132.11 -1.77H136.37Q136.77 -0.95 137.36 -0.60Q137.95 -0.24 138.51 -0.24Q138.51 -0.24 138.61 -0.24Q138.70 -0.24 138.70 -0.24V0.00H130.97V-0.24H131.19Q131.66 -0.24 132.04 -0.64Q132.41 -1.04 132.11 -1.77ZM153.58 -17.00V0.00H149.94V-17.00ZM159.99 -1.05 160.58 0.00H153.51V-1.05ZM161.77 -4.97 160.65 0.00H155.88L156.76 -1.05Q157.97 -1.05 158.90 -1.52Q159.83 -2.00 160.49 -2.88Q161.15 -3.76 161.53 -4.97ZM150.01 -1.77V0.00H148.07V-0.24Q148.07 -0.24 148.22 -0.24Q148.38 -0.24 148.38 -0.24Q149.01 -0.24 149.46 -0.69Q149.91 -1.14 149.94 -1.77ZM150.01 -15.23H149.94Q149.91 -15.86 149.46 -16.31Q149.01 -16.76 148.38 -16.76Q148.38 -16.76 148.22 -16.76Q148.07 -16.76 148.07 -16.76V-17.00H150.01ZM153.51 -15.23V-17.00H155.45V-16.76Q155.42 -16.76 155.28 -16.76Q155.13 -16.76 155.13 -16.76Q154.50 -16.76 154.05 -16.31Q153.60 -15.86 153.58 -15.23ZM204.51 -6.80V-2.79Q203.94 -2.06 202.78 -1.34Q201.63 -0.61 200.02 -0.13Q198.41 0.34 196.45 0.34Q193.70 0.33 191.63 -0.75Q189.57 -1.82 188.42 -3.81Q187.27 -5.79 187.27 -8.53Q187.27 -11.24 188.41 -13.21Q189.54 -15.19 191.64 -16.26Q193.74 -17.34 196.60 -17.34Q197.93 -17.34 199.20 -17.16Q200.47 -16.98 201.54 -16.68Q202.60 -16.39 203.31 -16.01L203.35 -11.85H203.13Q202.68 -13.51 201.81 -14.48Q200.93 -15.46 199.78 -15.89Q198.63 -16.32 197.33 -16.32Q195.35 -16.32 193.99 -15.36Q192.63 -14.39 191.93 -12.62Q191.23 -10.85 191.23 -8.38Q191.23 -6.04 191.88 -4.31Q192.52 -2.58 193.76 -1.63Q195.00 -0.69 196.77 -0.68Q197.68 -0.68 198.51 -0.96Q199.34 -1.24 199.97 -1.78Q200.60 -2.32 200.87 -3.09L200.89 -6.80Q200.89 -7.89 199.65 -7.89H199.26V-8.14H206.09V-7.89H205.72Q204.46 -7.89 204.51 -6.80ZM221.83 -17.00H224.96Q226.05 -17.00 227.09 -16.76Q228.13 -16.52 228.94 -15.98Q229.76 -15.44 230.24 -14.57Q230.72 -13.69 230.72 -12.42Q230.72 -11.40 230.30 -10.38Q229.88 -9.37 229.01 -8.66Q228.13 -7.94 226.74 -7.80Q227.52 -7.55 228.17 -6.92Q228.82 -6.29 229.23 -5.65Q229.26 -5.61 229.50 -5.22Q229.75 -4.82 230.13 -4.25Q230.51 -3.69 230.92 -3.10Q231.32 -2.52 231.67 -2.09Q232.20 -1.42 232.65 -1.02Q233.09 -0.63 233.57 -0.44Q234.05 -0.26 234.68 -0.24V0.00H232.31Q230.99 0.00 229.92 -0.24Q228.85 -0.47 228.03 -1.02Q227.21 -1.57 226.60 -2.52Q226.41 -2.82 226.18 -3.27Q225.95 -3.72 225.70 -4.22Q225.45 -4.72 225.23 -5.20Q225.00 -5.68 224.83 -6.07Q224.67 -6.46 224.59 -6.68Q224.26 -7.45 223.83 -7.84Q223.41 -8.23 222.94 -8.31V-8.55Q223.00 -8.55 223.22 -8.54Q223.45 -8.54 223.68 -8.55Q224.35 -8.56 224.97 -8.82Q225.59 -9.08 226.06 -9.68Q226.54 -10.29 226.75 -11.35Q226.79 -11.58 226.83 -11.87Q226.86 -12.17 226.84 -12.53Q226.77 -14.27 225.97 -15.09Q225.17 -15.90 223.97 -15.93Q223.63 -15.95 223.23 -15.94Q222.83 -15.94 222.51 -15.94Q222.19 -15.94 222.07 -15.94Q222.07 -15.95 222.01 -16.21Q221.95 -16.47 221.89 -16.74Q221.83 -17.00 221.83 -17.00ZM222.14 -17.00V0.00H218.48V-17.00ZM218.56 -1.77V0.00H216.61V-0.24Q216.64 -0.24 216.78 -0.24Q216.92 -0.24 216.93 -0.24Q217.56 -0.24 218.01 -0.69Q218.46 -1.14 218.48 -1.77ZM218.56 -15.23H218.48Q218.46 -15.86 218.01 -16.31Q217.56 -16.76 216.93 -16.76Q216.92 -16.76 216.78 -16.76Q216.64 -16.76 216.61 -16.76V-17.00H218.56ZM222.07 -1.77H222.14Q222.14 -1.14 222.60 -0.69Q223.06 -0.24 223.70 -0.24Q223.73 -0.24 223.85 -0.24Q223.97 -0.24 223.99 -0.24V0.00H222.07ZM253.43 -17.34Q256.36 -17.34 258.51 -16.26Q260.66 -15.19 261.83 -13.21Q263.00 -11.24 263.00 -8.50Q263.00 -5.78 261.83 -3.80Q260.66 -1.81 258.51 -0.74Q256.36 0.34 253.43 0.34Q250.52 0.34 248.38 -0.74Q246.23 -1.81 245.06 -3.79Q243.89 -5.76 243.89 -8.50Q243.89 -11.22 245.06 -13.20Q246.23 -15.19 248.38 -16.26Q250.52 -17.34 253.43 -17.34ZM253.43 -0.68Q255.16 -0.68 256.41 -1.64Q257.65 -2.59 258.33 -4.34Q259.01 -6.09 259.01 -8.50Q259.01 -10.91 258.33 -12.66Q257.65 -14.41 256.41 -15.36Q255.16 -16.32 253.43 -16.32Q251.73 -16.32 250.49 -15.36Q249.24 -14.41 248.56 -12.66Q247.88 -10.91 247.88 -8.50Q247.88 -6.09 248.56 -4.34Q249.24 -2.59 250.49 -1.64Q251.73 -0.68 253.43 -0.68ZM279.33 -17.00V-6.31Q279.33 -4.60 279.91 -3.33Q280.48 -2.07 281.54 -1.39Q282.59 -0.71 284.04 -0.71Q285.53 -0.71 286.60 -1.35Q287.68 -2.00 288.27 -3.18Q288.86 -4.37 288.86 -6.00V-17.00H290.59V-6.15Q290.59 -4.16 289.76 -2.71Q288.92 -1.25 287.35 -0.45Q285.78 0.34 283.59 0.34Q281.00 0.34 279.23 -0.45Q277.47 -1.24 276.58 -2.69Q275.69 -4.14 275.69 -6.12V-17.00ZM275.76 -17.00V-15.54H275.69Q275.69 -16.08 275.31 -16.42Q274.93 -16.76 274.40 -16.76Q274.40 -16.76 274.21 -16.76Q274.01 -16.76 274.01 -16.76V-17.00ZM281.01 -17.00V-16.76Q281.01 -16.76 280.81 -16.76Q280.62 -16.76 280.62 -16.76Q280.08 -16.76 279.71 -16.42Q279.33 -16.08 279.33 -15.54H279.28V-17.00ZM288.93 -17.00V-15.54H288.86Q288.86 -16.08 288.49 -16.42Q288.11 -16.76 287.57 -16.76Q287.57 -16.76 287.38 -16.76Q287.19 -16.76 287.19 -16.76V-17.00ZM292.27 -17.00V-16.76Q292.27 -16.76 292.07 -16.76Q291.88 -16.76 291.88 -16.76Q291.35 -16.76 290.97 -16.42Q290.59 -16.08 290.59 -15.54H290.54V-17.00ZM308.35 -17.00H311.26Q313.35 -17.00 314.66 -16.39Q315.97 -15.78 316.61 -14.73Q317.26 -13.68 317.30 -12.34Q317.34 -11.09 316.88 -10.08Q316.43 -9.07 315.62 -8.39Q314.81 -7.71 313.78 -7.40Q312.75 -7.10 311.65 -7.25Q310.54 -7.41 309.50 -8.09V-8.33Q309.50 -8.33 309.80 -8.33Q310.10 -8.32 310.57 -8.43Q311.05 -8.53 311.57 -8.80Q312.10 -9.07 312.56 -9.63Q313.03 -10.19 313.27 -11.10Q313.36 -11.42 313.40 -11.82Q313.44 -12.21 313.41 -12.56Q313.38 -14.11 312.56 -15.05Q311.75 -15.98 310.28 -15.98H308.60Q308.60 -15.98 308.54 -16.23Q308.48 -16.49 308.41 -16.74Q308.35 -17.00 308.35 -17.00ZM308.67 -17.00V0.00H305.03V-17.00ZM305.10 -1.77V0.00H303.16V-0.24Q303.19 -0.24 303.32 -0.24Q303.46 -0.24 303.47 -0.24Q304.10 -0.24 304.55 -0.69Q305.00 -1.14 305.03 -1.77ZM305.10 -15.23H305.03Q305.03 -15.86 304.57 -16.31Q304.10 -16.76 303.47 -16.76Q303.46 -16.76 303.33 -16.76Q303.21 -16.76 303.18 -16.76L303.16 -17.00H305.10ZM308.60 -1.77H308.67Q308.69 -1.14 309.14 -0.69Q309.59 -0.24 310.22 -0.24Q310.25 -0.24 310.39 -0.24Q310.53 -0.24 310.54 -0.24V0.00H308.60Z"/></g>
</svg>
'@
    return "<div class='logo'>$lockup</div>"
}

# ---------------------------------------------------------------------------
# Small inline PASS / FAIL / NA text badge (for Email Health table)
# ---------------------------------------------------------------------------
function Get-CellIcon {
    param([string]$Status, [int]$Size = 20)
    $s = ("$Status").Trim().ToUpper()
    switch ($s) {
        'PASS' { return "<span class='badge-pass'>PASS</span>" }
        'FAIL' { return "<span class='badge-fail'>FAIL</span>" }
        'NA'   { return "<span class='badge-na'>NA</span>" }
        default { return "<span class='badge-na'>$s</span>" }
    }
}

# ---------------------------------------------------------------------------
# Radar / spider chart (Microsoft Security Baseline)
#   $Axes = ordered array of @{ Label='...'; Value=<0..Max> }
# ---------------------------------------------------------------------------
function Get-RadarSvg {
    param($Axes, [double]$Max = 100)
    $n = @($Axes).Count
    if ($n -lt 3) { return "<svg width='420' height='420'></svg>" }
    $W = 460; $H = 460; $cx = 230; $cy = 235; $R = 150
    if ($Max -le 0) { $Max = 1 }
    $rings = 4
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg class='radar' width='$W' height='$H' viewBox='0 0 $W $H'>")
    # grid rings (polygons)
    for ($ring = 1; $ring -le $rings; $ring++) {
        $rr = $R * $ring / $rings
        $pts = @()
        for ($i = 0; $i -lt $n; $i++) {
            $ang = (-90 + 360.0 * $i / $n) * [math]::PI / 180.0
            $x = [math]::Round($cx + $rr * [math]::Cos($ang), 1)
            $y = [math]::Round($cy + $rr * [math]::Sin($ang), 1)
            $pts += "$x,$y"
        }
        [void]$sb.Append("<polygon points='$($pts -join ' ')' fill='none' stroke='$COLOR_BORDER' stroke-width='1'/>")
    }
    # spokes + labels
    for ($i = 0; $i -lt $n; $i++) {
        $ang = (-90 + 360.0 * $i / $n) * [math]::PI / 180.0
        $x = [math]::Round($cx + $R * [math]::Cos($ang), 1)
        $y = [math]::Round($cy + $R * [math]::Sin($ang), 1)
        [void]$sb.Append("<line x1='$cx' y1='$cy' x2='$x' y2='$y' stroke='$COLOR_BORDER' stroke-width='1'/>")
        $lx = [math]::Round($cx + ($R + 26) * [math]::Cos($ang), 1)
        $ly = [math]::Round($cy + ($R + 26) * [math]::Sin($ang), 1)
        $anchor = 'middle'
        if ($lx -gt $cx + 10) { $anchor = 'start' } elseif ($lx -lt $cx - 10) { $anchor = 'end' }
        $lbl = Enc ($Axes[$i].Label)
        [void]$sb.Append("<text x='$lx' y='$($ly+4)' text-anchor='$anchor' fill='$COLOR_BODY' font-size='12'>$lbl</text>")
    }
    # data polygon
    $dpts = @()
    for ($i = 0; $i -lt $n; $i++) {
        $v = [double]$Axes[$i].Value
        if ($v -lt 0) { $v = 0 }
        if ($v -gt $Max) { $v = $Max }
        $rr = $R * ($v / $Max)
        $ang = (-90 + 360.0 * $i / $n) * [math]::PI / 180.0
        $x = [math]::Round($cx + $rr * [math]::Cos($ang), 1)
        $y = [math]::Round($cy + $rr * [math]::Sin($ang), 1)
        $dpts += "$x,$y"
    }
    [void]$sb.Append("<polygon points='$($dpts -join ' ')' fill='$COLOR_GOLD' fill-opacity='0.28' stroke='$COLOR_GOLD' stroke-width='2.5'/>")
    foreach ($p in $dpts) {
        $xy = $p -split ','
        [void]$sb.Append("<circle cx='$($xy[0])' cy='$($xy[1])' r='3.2' fill='$COLOR_GOLD'/>")
    }
    [void]$sb.Append('</svg>')
    return $sb.ToString()
}

# ---------------------------------------------------------------------------
# Half-circle gauge (Security Baseline Overview tiers)
# ---------------------------------------------------------------------------
function Get-HalfGaugeSvg {
    param([int]$Value, [int]$Total, [string]$Label)
    $W = 200; $H = 128; $cx = 100; $cy = 110; $r = 82
    $pct = 0.0
    if ($Total -gt 0) { $pct = [math]::Max(0, [math]::Min(1, $Value / [double]$Total)) }
    # semicircle path from left (180deg) to right (0deg)
    $circ = [math]::PI * $r            # length of semicircle
    $fill = [math]::Round($circ * $pct, 2)
    $rest = [math]::Round($circ - $fill, 2)
    $circR = [math]::Round($circ, 2)
    return @"
<svg class='halfgauge' width='$W' height='$H' viewBox='0 0 $W $H'>
  <path d='M $($cx-$r) $cy A $r $r 0 0 1 $($cx+$r) $cy' fill='none' stroke='$COLOR_BORDER' stroke-width='14' stroke-linecap='round'
        stroke-dasharray='$circR $circR'/>
  <path d='M $($cx-$r) $cy A $r $r 0 0 1 $($cx+$r) $cy' fill='none' stroke='$COLOR_GOLD' stroke-width='14' stroke-linecap='round'
        stroke-dasharray='$fill $rest'/>
  <text x='$cx' y='$($cy-6)' text-anchor='middle' fill='$COLOR_TEXT' font-size='30' font-weight='700'>$Value/$Total</text>
  <text x='$cx' y='$($cy+16)' text-anchor='middle' fill='$COLOR_GOLD' font-size='13' font-weight='600' letter-spacing='1'>$(Enc $Label)</text>
</svg>
"@
}

# ===========================================================================
# EARLY AGGREGATES - computed once, up front, so the Executive Summary and
# Recommendations pages can reference them before their "home" pages
# (Secure Score, User Health, the per-workload check pages) build their own
# fuller presentation of the same underlying data. Single source of truth:
# neither page recomputes these.
# ===========================================================================
$ss = $data.SecureScore
$ssPct = 0; $ssCur = 0.0; $ssMax = 0.0
if ($ss) {
    try { $ssCur = [double]$ss.Current } catch {}
    try { $ssMax = [double]$ss.Max } catch {}
    if ($ss.Percentage -ne $null) { try { $ssPct = [int]$ss.Percentage } catch {} }
    elseif ($ssMax -gt 0) { $ssPct = [int][math]::Round(($ssCur/$ssMax)*100) }
}

$uh = $data.UserHealth
$totalUsers = [int]($uh.TotalUsers)
$usersWithoutMfa = [int]($uh.UsersWithoutMFA)

$controls = @($data.Controls)

# ---------------------------------------------------------------------------
# Shared: parent status per CheckId (FAIL > WARN > MANUAL > PASS aggregation)
# ---------------------------------------------------------------------------
function Get-ParentStatus {
    param($Items)
    $statuses = @($Items | ForEach-Object { ("$($_.Status)").Trim().ToUpper() })
    $hasFail = ($statuses -contains 'FAIL')
    $allManual = ($statuses.Count -gt 0) -and -not ($statuses | Where-Object { $_ -notmatch '^(MANUAL|NOTSET|NOT SET|NA|N/A)$' })
    $allPass = ($statuses.Count -gt 0) -and -not ($statuses | Where-Object { $_ -ne 'PASS' })
    if ($hasFail) { return 'FAIL' }
    elseif ($allPass) { return 'PASS' }
    elseif ($allManual) { return 'MANUAL' }
    else { return 'WARN' }
}
$parentStatusMap = @{}
foreach ($grp in ($controls | Group-Object CheckId)) {
    $parentStatusMap[[string]$grp.Name] = Get-ParentStatus -Items @($grp.Group)
}

# Fail counts by severity, for the Executive Summary stat cards and to
# decide the headline risk framing.
$failedControls = @($controls | Where-Object { ("$($_.Status)").Trim().ToUpper() -eq 'FAIL' })
$critFailCount = @($failedControls | Where-Object { ("$($_.Severity)").Trim().ToLower() -eq 'critical' }).Count
$highFailCount = @($failedControls | Where-Object { ("$($_.Severity)").Trim().ToLower() -eq 'high' }).Count
$totalFailCount = $failedControls.Count
$totalControlCount = @($controls | Select-Object -Unique CheckId).Count

# ===========================================================================
# BUILD PAGES
# ===========================================================================
$pages = New-Object System.Text.StringBuilder

# ---- Page 1: Cover -------------------------------------------------------
$logoHtml = Get-LogoHtml
# Tier badge: the first thing a reader sees, so the free self-serve
# assessment and the paid Advanced (certificate-based) engagement are never
# mistaken for each other - Scott's explicit ask after reviewing a real
# Advanced-tier report that read like the free one throughout.
$tierBadgeLabel = if ($IsAdvancedTier) { 'Advanced Assessment' } else { 'Free Assessment' }
[void]$pages.Append(@"
<section class='page cover'>
  $Watermark
  <div class='cover-top'>$logoHtml</div>
  <div class='cover-body'>
    <div class='cover-date'>$(Enc $reportDateStr)</div>
    <h1 class='cover-name'>$(Enc $CustomerName)</h1>
    <h2 class='cover-title'>Cloud Assessment Report</h2>
    <div class='cover-tier-badge'>$(Enc $tierBadgeLabel)</div>
    <div class='cover-rule'></div>
    <div class='cover-exec'>Executive Summary</div>
  </div>
  <div class='cover-foot'>Prepared by $(Enc $PreparedBy)</div>
</section>
"@)

# ---- Page 2: Executive Summary -------------------------------------------
function Get-RiskBand {
    param([int]$Pct)
    if ($Pct -ge 80) { return @{ Label = 'a strong security posture'; Note = 'Most of the Microsoft-recommended controls we checked are already in place.' } }
    elseif ($Pct -ge 60) { return @{ Label = 'a solid foundation with clear room to strengthen it'; Note = 'The core controls are largely in place; a handful of specific gaps stand out below.' } }
    elseif ($Pct -ge 40) { return @{ Label = 'moderate exposure that is worth addressing soon'; Note = 'Several important controls are missing or only partially configured.' } }
    else { return @{ Label = 'significant exposure that warrants prompt action'; Note = 'A number of foundational controls are not yet in place.' } }
}
$riskBand = Get-RiskBand -Pct $ssPct
$mfaGapNote = if ($totalUsers -gt 0 -and $usersWithoutMfa -gt 0) {
    $mfaPct = [math]::Round(100.0 * $usersWithoutMfa / $totalUsers, 0)
    " $usersWithoutMfa of $('{0:N0}' -f $totalUsers) users ($mfaPct%) do not have multi-factor authentication enabled - typically the single highest-impact gap to close first."
} elseif ($totalUsers -gt 0) {
    ' Every user in the tenant has multi-factor authentication enabled, which is one of the most effective controls available.'
} else { '' }

[void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='page-h1'>Executive Summary</h1>
  <p class='exec-summary'>
    $(Enc $CustomerName)'s Microsoft 365 environment currently scores <strong>$ssPct%</strong> against Microsoft's Secure Score benchmark, reflecting <strong>$($riskBand.Label)</strong>. $($riskBand.Note) Out of $totalControlCount security controls assessed across Entra ID, Exchange, Teams, Intune, SharePoint/OneDrive, Defender and Purview, <strong>$totalFailCount</strong> did not pass - including <strong>$critFailCount critical</strong> and <strong>$highFailCount high-severity</strong> findings.$(Enc $mfaGapNote)
  </p>
  <p class='exec-summary'>
    The pages that follow set out exactly what we checked and why, a prioritised list of what to fix first, and the full detail behind every control. This is a point-in-time, read-only assessment - see <strong>Scope &amp; Methodology</strong> overleaf for what it does and does not cover.
  </p>
  <div class='exec-stat-row'>
    <div class='exec-stat'><div class='exec-stat-num'>$ssPct%</div><div class='exec-stat-label'>Secure Score</div></div>
    <div class='exec-stat'><div class='exec-stat-num'>$totalControlCount</div><div class='exec-stat-label'>Controls Assessed</div></div>
    <div class='exec-stat'><div class='exec-stat-num'>$critFailCount</div><div class='exec-stat-label'>Critical Findings</div></div>
    <div class='exec-stat'><div class='exec-stat-num'>$('{0:N0}' -f $totalUsers)</div><div class='exec-stat-label'>Users in Tenant</div></div>
  </div>
</section>
"@)

# ---- Page 3: Scope & Methodology ------------------------------------------
# Tier-specific copy: an earlier version of this page always described the
# free, self-serve assessment regardless of which tier actually ran -
# confirmed live, a real Advanced-tier report read like the free one
# throughout. Access method and workload depth genuinely differ by tier
# (client-secret/Graph-only vs. certificate-based with full Exchange/Teams
# coverage), so the copy below does too.
if ($IsAdvancedTier) {
    $scopeIntro = "This is Camelot Digital Group's Advanced assessment, conducted as part of our active engagement with you - a deeper, certificate-based review that reaches further than our free self-serve assessment can. Here is exactly what that means in practice."
    $scopeItems = @(
        'We reviewed Entra ID (identity, MFA, conditional access, privileged roles), Exchange Online (full mailbox security, email authentication and anti-phishing), Teams (external access, meeting and messaging policies), Intune (device compliance and management), SharePoint/OneDrive (external sharing), Microsoft Defender, and Purview (retention and data loss prevention).'
        'Access was read-only, granted via a certificate-based Microsoft Entra app registration with additional Exchange Online and Microsoft Teams permissions beyond the free assessment - the same permission model Microsoft itself recommends for security reviews. No configuration was changed at any point.'
        'The scripts that performed this assessment are open source and published on our public GitHub, so you (or your own IT team) can review exactly what was run before, during, or after granting access.'
        'This reflects a single point in time. Microsoft 365 configurations change as people, policies and licensing change, so we would expect a re-assessment some months from now to look different.'
    )
    $scopeNoteClose = 'If any of these matter to you, let us know and we can scope them as part of this engagement - see the closing page for how to reach us.'
} else {
    $scopeIntro = 'This is a free, self-serve assessment designed to give you a fast, honest read on your Microsoft 365 security posture - not a substitute for a full audit. Here is exactly what that means in practice.'
    $scopeItems = @(
        'We reviewed Entra ID (identity, MFA, conditional access, privileged roles), Exchange Online (email authentication and anti-phishing), Teams, Intune (device compliance and management), SharePoint/OneDrive (external sharing), Microsoft Defender, and Purview (retention and data loss prevention).'
        'Access was read-only, granted via a multi-tenant Microsoft Entra app registration using the Microsoft Graph API - the same permission model Microsoft itself recommends for security reviews. No configuration was changed at any point.'
        'The scripts that performed this assessment are open source and published on our public GitHub, so you (or your own IT team) can review exactly what was run before, during, or after granting access.'
        'This reflects a single point in time. Microsoft 365 configurations change as people, policies and licensing change, so we would expect a re-assessment some months from now to look different.'
    )
    $scopeNoteClose = 'If any of these matter to you, our team can scope a deeper engagement - see the closing page for how to reach us.'
}
[void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='page-h1'>Scope &amp; Methodology</h1>
  <p class='body-text'>$scopeIntro</p>
  <div class='scope-list'>
    $(($scopeItems | ForEach-Object { "<div class='scope-item'><div class='scope-mark'></div><div class='scope-text'>$_</div></div>" }) -join "`n")
  </div>
  <div class='scope-note'>
    <strong>What this assessment does not include:</strong> penetration testing, phishing simulations, physical security review, staff interviews, or a formal compliance certification (e.g. Cyber Essentials, ISO 27001, SOC 2). It also cannot see configuration outside Microsoft 365 - network security, third-party SaaS, or on-premises systems are out of scope. $scopeNoteClose
  </div>
</section>
"@)

# ---- Page 4: Secure Score (moved after Executive Summary / Scope pages) --
# $ss/$ssPct/$ssCur/$ssMax computed early (see EARLY AGGREGATES) so the
# Executive Summary page can reference them too.
$gauge = Get-GaugeSvg -Percentage $ssPct -Current $ssCur -Max $ssMax
$chart = Get-LineChartSvg -History ($ss.History)
$tcs = [math]::Max(0, [math]::Min(100, $TypicalClientScore))
$yourPos = [math]::Max(0, [math]::Min(100, $ssPct))

[void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='page-h1'>Secure Score</h1>
  <div class='ss-row'>
    <div class='ss-gauge'>$gauge</div>
    <div class='ss-chart'>$chart</div>
  </div>
  <p class='body-text'>The Secure Score is a reflection of your organization's security posture. It is a measure of how well your organization is leveraging the security features in Microsoft 365. The Secure Score is calculated based on the security features that you have enabled and the actions that you have taken to protect your organization. The higher the score, the more secure your organization is.</p>
  <div class='cmp-wrap'>
    <div class='cmp-label-left'>High<br>Risk</div>
    <div class='cmp-track-wrap'>
      <div class='cmp-marker-top' style='left:$yourPos%;'><span class='cmp-tip'>Your Business: $ssPct%</span></div>
      <div class='cmp-track'>
        <div class='cmp-dot' style='left:$yourPos%;'></div>
        <div class='cmp-dot' style='left:$tcs%;'></div>
      </div>
      <div class='cmp-marker-bot' style='left:$tcs%;'><span class='cmp-tip'>Our Typical Client: $tcs%</span></div>
    </div>
    <div class='cmp-label-right'>Low<br>Risk</div>
  </div>
</section>
"@)

# ---- Page 5: User Health (moved after Executive Summary / Scope pages) --
# $uh/$totalUsers computed early (see EARLY AGGREGATES).
function _lvlUsersNoMfa($v,$total) {
    if ($v -le 0) { return 'ok' }
    if ($total -gt 0 -and (($v / [double]$total) -lt 0.05)) { return 'warn' }
    return 'bad'
}
$rowsUser = @(
    @{ icon='person'; num=$totalUsers; lvl='ok';
       label='Total Users';
       desc='The total number of users in the tenant. This includes all users registered in Entra including unlicensed users, guest users, and service accounts.' }
    @{ icon='admin'; num=[int]$uh.GlobalAdmins;
       lvl=$(if([int]$uh.GlobalAdmins -ge 2 -and [int]$uh.GlobalAdmins -le 4){'ok'}else{'warn'});
       label='Tenants should have 2-4 users with the Global Administrator role';
       desc='Global Administrators have full access to all administrative features in the tenant. It is recommended to have between two and four global administrators. Excessive global administrators increase the risk of unauthorized access to the tenant.' }
    @{ icon='lock'; num=[int]$uh.UsersWithoutMFA;
       lvl=(_lvlUsersNoMfa ([int]$uh.UsersWithoutMFA) $totalUsers);
       label='Users without Multi-Factor Authentication';
       desc='Multi-Factor Authentication (MFA) requires users to provide two or more verification factors to sign in. Users without MFA are at a higher risk of unauthorized access to their account.' }
    @{ icon='lock'; num=[int]$uh.UsersWithWeakMFA;
       lvl=$(if([int]$uh.UsersWithWeakMFA -le 0){'ok'}else{'info'});
       label='Users with weak Multi-Factor Authentication';
       desc='Users with weak MFA have MFA enabled, but are using weak authentication methods such as SMS, Voice, and Email. These methods are less secure and can be more easily compromised.' }
    @{ icon='risky'; num=[int]$uh.RiskyUsers;
       lvl=$(if([int]$uh.RiskyUsers -le 0){'ok'}else{'bad'});
       label='Users with risky sign-ins';
       desc="Risky users are users who have had risky sign-ins. Risky sign-ins can indicate that a user's account has been compromised or is at risk of being compromised. It is important to review risky users and take action to secure their accounts." }
    @{ icon='person'; num=[int]$uh.DormantUsers;
       lvl=$(if([int]$uh.DormantUsers -le 0){'ok'}else{'warn'});
       label='Dormant / inactive users';
       desc='Dormant users have not signed in for an extended period. Inactive accounts increase the attack surface and should be reviewed and disabled if no longer required.' }
)
$userRowsHtml = ($rowsUser | ForEach-Object {
    $ind = Get-StatIndicator -Level $_.lvl
    $mic = Get-MetricIcon -Kind $_.icon
    "<div class='metric-row'><div class='metric-circle'>$mic</div><div class='metric-mid'><div class='metric-num'>$('{0:N0}' -f $_.num)<span class='metric-ind'>$ind</span></div><div class='metric-label'>$(Enc $_.label)</div></div><div class='metric-desc'>$(Enc $_.desc)</div></div>"
}) -join "`n"

[void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='page-h1'>User Health</h1>
  <div class='metric-list'>
    $userRowsHtml
  </div>
</section>
"@)

# ---- Page 4: Device Health -----------------------------------------------
$dh = $data.DeviceHealth
$rowsDev = @(
    @{ icon='devices'; num=[int]$dh.EntraDevices;
       lvl=$(if([int]$dh.EntraDevices -gt 0){'ok'}else{'na'});
       label='Devices enrolled in Microsoft Entra';
       desc='Entra is a device management solution that provides a single pane of glass for managing devices across multiple platforms.' }
    @{ icon='lock-open'; num=[int]$dh.NotEncryptedDevices;
       lvl=$(if([int]$dh.NotEncryptedDevices -le 0){'ok'}elseif([int]$dh.NotEncryptedDevices -lt 5){'info'}else{'bad'});
       label='Devices without encryption enabled';
       desc='Devices without encryption enabled are at risk of data exposure due to theft or loss.' }
    @{ icon='doc-alert'; num=[int]$dh.NonCompliantDevices;
       lvl=$(if([int]$dh.NonCompliantDevices -le 0){'ok'}else{'bad'});
       label="Devices that are not compliant with the organization's security policies";
       desc="Devices that are not compliant with the organization's security policies are at risk of being compromised and should be investigated immediately." }
    @{ icon='calendar'; num=[int]$dh.StaleDevices;
       lvl=$(if([int]$dh.StaleDevices -le 0){'ok'}else{'warn'});
       label='Devices that have not been used in the last 30 days';
       desc='Stale devices are at greater risk of being compromised due to lack of security updates and patches and potential loss or theft.' }
)
$devRowsHtml = ($rowsDev | ForEach-Object {
    $ind = Get-StatIndicator -Level $_.lvl
    $mic = Get-MetricIcon -Kind $_.icon
    "<div class='metric-row'><div class='metric-circle'>$mic</div><div class='metric-mid'><div class='metric-num'>$('{0:N0}' -f $_.num)<span class='metric-ind'>$ind</span></div><div class='metric-label'>$(Enc $_.label)</div></div><div class='metric-desc'>$(Enc $_.desc)</div></div>"
}) -join "`n"

[void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='page-h1'>Device Health</h1>
  <div class='metric-list'>
    $devRowsHtml
  </div>
</section>
"@)

# ---- Page 5: Applications & Data -----------------------------------------
$ad = $data.ApplicationsData
$shareVal = [string]$ad.SPSharingSetting
$shareLabel = Get-SharingLabel -Value $shareVal
$shareDesc  = Get-SharingDescription -Value $shareVal
$shareStat  = Get-SharingSeverity -Value $shareVal
$shareIcon  = Get-StatusIcon -Status $shareStat -Size 40

$spRowsHtml = ''
if ($ad.SPSitesPublicLinks) {
    $spSorted = $ad.SPSitesPublicLinks | Sort-Object { [int]$_.PublicLinkCount } -Descending | Select-Object -First 15
    $spRowsHtml = ($spSorted | ForEach-Object {
        "<div class='sp-row'><div class='sp-name'>$(Enc $_.SiteDisplayName)</div><div class='sp-count'><div class='sp-count-num'>$([int]$_.PublicLinkCount)</div><div class='sp-count-lbl'>Public Count</div></div></div>"
    }) -join "`n"
} else {
    $spRowsHtml = "<div class='sp-empty'>No SharePoint sites with public links were detected.</div>"
}
$enterpriseApps = 0
if ($ad.EnterpriseApps -ne $null) { try { $enterpriseApps = [int]$ad.EnterpriseApps } catch {} }

# --- SharePoint site inventory (all sites) ---------------------------------
$spSites = @()
if ($ad.PSObject.Properties.Name -contains 'SPSites' -and $ad.SPSites) {
    $spSites = @($ad.SPSites)
}
$spTotalSites = 0
if ($ad.PSObject.Properties.Name -contains 'SPTotalSites' -and $ad.SPTotalSites -ne $null) {
    try { $spTotalSites = [int]$ad.SPTotalSites } catch {}
}
if ($spTotalSites -eq 0) { $spTotalSites = $spSites.Count }
$spSitesLine = if ($spTotalSites -gt 0) {
    "$('{0:N0}' -f $spTotalSites) SharePoint site collection(s) were discovered in the tenant. The full inventory is listed below."
} else {
    "No SharePoint sites were retrieved. Verify the app has the <b>Sites.Read.All</b> application permission with admin consent granted."
}

[void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='page-h1'>Applications &amp; Data</h1>
  <div class='ad-sub'>Enterprise Applications</div>
  <p class='body-text'>$('{0:N0}' -f $enterpriseApps) enterprise applications are registered in the tenant. Applications should be cataloged and periodically reviewed to ensure only approved apps retain access.</p>
  <div class='ad-sub'>Default Sharing Policy</div>
  <div class='share-row'>
    <div class='share-icon'>$shareIcon</div>
    <div class='share-body'>
      <div class='share-title'>$(Enc $shareLabel)</div>
      <div class='share-desc'>$(Enc $shareDesc)</div>
    </div>
  </div>
  <div class='ad-sub'>Top SharePoint Sites Public Links</div>
  <div class='sp-list'>
    $spRowsHtml
  </div>
</section>
"@)

# ===========================================================================
# SharePoint Site Inventory (paginated) - lists ALL discovered sites
# ===========================================================================
if ($spSites.Count -gt 0) {
    $perPage = 22
    $chunks = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $spSites.Count; $i += $perPage) {
        $end = [Math]::Min($i + $perPage, $spSites.Count) - 1
        $chunks.Add(@($spSites[$i..$end]))
    }
    $pageNo = 0
    foreach ($chunk in $chunks) {
        $pageNo++
        $rows = ($chunk | ForEach-Object {
            $nm = [string]$_.DisplayName
            if ([string]::IsNullOrWhiteSpace($nm)) { $nm = [string]$_.Name }
            if ([string]::IsNullOrWhiteSpace($nm)) { $nm = '(unnamed site)' }
            $url = [string]$_.WebUrl
            "<tr><td class='spi-name'>$(Enc $nm)</td><td class='spi-url'>$(Enc $url)</td></tr>"
        }) -join "`n"
        $subhdr = if ($chunks.Count -gt 1) { " <span class='spi-page'>(page $pageNo of $($chunks.Count))</span>" } else { '' }
        [void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='page-h1'>SharePoint Site Inventory$subhdr</h1>
  <p class='body-text'>$spSitesLine</p>
  <table class='spi-table'>
    <thead><tr><th>Site Name</th><th>URL</th></tr></thead>
    <tbody>
      $rows
    </tbody>
  </table>
</section>
"@)
    }
}

# $controls / Get-ParentStatus / $parentStatusMap computed early (see
# EARLY AGGREGATES) so the Recommendations page can use them too.

# ===========================================================================
# Page 6: Email Health
# ===========================================================================
$emailHealth = @()
if ($data.PSObject.Properties.Name -contains 'EmailHealth' -and $data.EmailHealth) {
    $emailHealth = @($data.EmailHealth)
}
$ehRowsHtml = ''
if ($emailHealth.Count -gt 0) {
    # default domain first, then onmicrosoft domains last, alpha within groups
    $ehSorted = $emailHealth | Sort-Object `
        @{ Expression = { if ([bool]$_.IsDefault) { 0 } else { 1 } } }, `
        @{ Expression = { if ([bool]$_.IsOnMicrosoft) { 1 } else { 0 } } }, `
        @{ Expression = { [string]$_.Domain } }
    $ehRowsHtml = ($ehSorted | ForEach-Object {
        $defBadge = if ([bool]$_.IsDefault) { "<span class='eh-default'>Yes</span>" } else { "<span class='eh-no'>No</span>" }
        $dmarcCell = Get-CellIcon -Status ([string]$_.Dmarc)
        $dmarcPol  = [string]$_.DmarcPolicy
        $dmarcPolHtml = if (-not [string]::IsNullOrWhiteSpace($dmarcPol)) { "<span class='eh-pol'>$(Enc $dmarcPol)</span>" } else { '' }
        "<tr>
           <td class='eh-dom'>$(Enc $_.Domain)</td>
           <td class='eh-c'>$defBadge</td>
           <td class='eh-c'>$(Get-CellIcon -Status ([string]$_.Spf))</td>
           <td class='eh-c'>$(Get-CellIcon -Status ([string]$_.Verified))</td>
           <td class='eh-c'>$(Get-CellIcon -Status ([string]$_.Dkim))</td>
           <td class='eh-c'>$dmarcCell $dmarcPolHtml</td>
         </tr>"
    }) -join "`n"
} else {
    $ehRowsHtml = "<tr><td colspan='6' class='eh-empty'>No domain email-authentication data was collected.</td></tr>"
}

# mail flow stats
$mfScanned = 0; $mfDelivered = 0; $mfBlocked = 0
if ($data.PSObject.Properties.Name -contains 'MailFlow' -and $data.MailFlow) {
    try { $mfScanned   = [int64]$data.MailFlow.EmailsScanned }   catch {}
    try { $mfDelivered = [int64]$data.MailFlow.EmailsDelivered } catch {}
    try { $mfBlocked   = [int64]$data.MailFlow.EmailsBlocked }   catch {}
}
$mfMax = [math]::Max($mfScanned, [math]::Max($mfDelivered, $mfBlocked))
if ($mfMax -le 0) { $mfMax = 1 }
$wScanned   = [math]::Round(100.0 * $mfScanned   / $mfMax, 1)
$wDelivered = [math]::Round(100.0 * $mfDelivered / $mfMax, 1)
$wBlocked   = [math]::Round(100.0 * $mfBlocked   / $mfMax, 1)
$mfBarsHtml = ''
if ($mfScanned -gt 0 -or $mfDelivered -gt 0 -or $mfBlocked -gt 0) {
    $mfBarsHtml = @"
<div class='mf-bars'>
  <div class='mf-item'>
    <div class='mf-top'><span class='mf-label'>Emails Scanned</span><span class='mf-num'>$('{0:N0}' -f $mfScanned)</span></div>
    <div class='mf-track'><div class='mf-fill' style='width:$wScanned%; background:$COLOR_GOLD;'></div></div>
  </div>
  <div class='mf-item'>
    <div class='mf-top'><span class='mf-label'>Emails Delivered</span><span class='mf-num'>$('{0:N0}' -f $mfDelivered)</span></div>
    <div class='mf-track'><div class='mf-fill' style='width:$wDelivered%; background:$COLOR_GREEN;'></div></div>
  </div>
  <div class='mf-item'>
    <div class='mf-top'><span class='mf-label'>Emails Blocked</span><span class='mf-num'>$('{0:N0}' -f $mfBlocked)</span></div>
    <div class='mf-track'><div class='mf-fill' style='width:$wBlocked%; background:$COLOR_RED;'></div></div>
  </div>
</div>
<p class='mf-note'>Mail-flow volume observed over the last 30 days.</p>
"@
} else {
    $mfBarsHtml = "<p class='mf-note'>No mail-flow statistics were available for the last 30 days.</p>"
}

[void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='page-h1'>Email Health</h1>
  <p class='body-text'>Email authentication protects your domains from spoofing and improves deliverability. The table below shows the status of SPF, DKIM and DMARC for each domain, along with whether the domain is verified for mail flow and its DMARC enforcement policy.</p>
  <table class='eh-table'>
    <thead>
      <tr>
        <th class='eh-dom'>Domain</th>
        <th>Default</th>
        <th>SPF</th>
        <th>Verified</th>
        <th>DKIM</th>
        <th>DMARC</th>
      </tr>
    </thead>
    <tbody>
      $ehRowsHtml
    </tbody>
  </table>
  <div class='ad-sub'>Mail Flow</div>
  $mfBarsHtml
</section>
"@)

# ===========================================================================
# Page 6b: Licensing Overview
# ===========================================================================
$licRowsHtml = ''
if ($data.PSObject.Properties.Name -contains 'Licensing' -and $data.Licensing) {
    $lics = @($data.Licensing)
    if ($lics.Count -gt 0) {
        $licRowsHtml = ($lics | Sort-Object { [int]$_.Total } -Descending | ForEach-Object {
            $consumed = [int]$_.Consumed
            $total    = [int]$_.Total
            $pct      = if ($total -gt 0) { [math]::Round(100.0 * $consumed / $total, 0) } else { 0 }
            # Single-quoted '$COLOR_RED'/'$COLOR_GREEN' don't interpolate in
            # PowerShell - that literally put the text "$COLOR_GREEN" into
            # the rendered CSS instead of the real hex colour. Confirmed
            # live: every licence bar rendered with no fill colour at all.
            $barColor = if ($pct -gt 90) { $COLOR_RED } elseif ($pct -gt 70) { '#f59e0b' } else { $COLOR_GREEN }
            $name     = [string]$_.SkuPartNumber
            # Make the SKU name more readable
            $friendlyName = ($name -replace '_', ' ')
            "<tr>
               <td class='lic-name'>$(Enc $friendlyName)</td>
               <td>$consumed</td>
               <td>$total</td>
               <td><div class='lic-bar-track'><div class='lic-bar-fill' style='width:$pct%; background:$barColor;'></div></div><span class='lic-pct'>$pct%</span></td>
             </tr>"
        }) -join "`n"
    }
}
if ([string]::IsNullOrWhiteSpace($licRowsHtml)) {
    $licRowsHtml = "<tr><td colspan='4' style='color:rgba(30,33,36,0.55);padding:12px;text-align:center;'>No licensing data was available. Ensure <code>Organization.Read.All</code> is granted.</td></tr>"
}

[void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='page-h1'>Licensing Overview</h1>
  <p class='body-text'>The table below shows the Microsoft 365 subscription licences assigned in the tenant, with consumption rates. Licences approaching full capacity are highlighted and may require expansion or review.</p>
  <table class='lic-table'>
    <thead>
      <tr>
        <th>Licence</th>
        <th>Consumed</th>
        <th>Total</th>
        <th>Utilisation</th>
      </tr>
    </thead>
    <tbody>
      $licRowsHtml
    </tbody>
  </table>
</section>
"@)

# ===========================================================================
# Page 6c: Authentication Methods & Guest Access
# ===========================================================================
$amCardsHtml = ''
if ($data.PSObject.Properties.Name -contains 'AuthMethods' -and $data.AuthMethods) {
    $methods = @($data.AuthMethods)
    if ($methods.Count -gt 0) {
        # Map internal method IDs to friendly names
        $friendlyNames = @{
            'microsoftAuthenticator'     = 'Authenticator App'
            'fido2'                      = 'FIDO2 Security Key'
            'sms'                        = 'SMS'
            'email'                      = 'Email OTP'
            'temporaryAccessPass'        = 'Temporary Access Pass'
            'softwareOath'               = 'Software OATH Token'
            'voice'                      = 'Voice Call'
            'x509Certificate'            = 'Certificate (x509)'
            'hardwareOath'               = 'Hardware OATH Token'
            'verifiableCredential'       = 'Verifiable Credentials'
            'qrCodePin'                  = 'QR Code + PIN'
        }
        $amCardsHtml = ($methods | ForEach-Object {
            $mid   = [string]$_.Method
            $state = [string]$_.State
            # -creplace (case-SENSITIVE), not -replace: PowerShell's -replace is
            # case-insensitive by default, which silently turns this
            # lowercase-then-uppercase boundary pattern into "any two adjacent
            # letters", inserting a space between every letter pair instead of
            # just at camelCase boundaries - confirmed live, produced
            # "V er if ia bl eC re de nt ia l" for an unmapped method ID.
            $fname = if ($friendlyNames.ContainsKey($mid)) { $friendlyNames[$mid] } else { $mid -creplace '([a-z])([A-Z])','$1 $2' }
            $cls   = switch ($state.ToLower()) { 'enabled' { 'am-enabled' }; 'disabled' { 'am-disabled' }; default { 'am-other' } }
            "<div class='am-card'><div class='am-name'>$(Enc $fname)</div><div class='am-state $cls'>$(Enc $state)</div></div>"
        }) -join "`n"
    }
}
if ([string]::IsNullOrWhiteSpace($amCardsHtml)) {
    $amCardsHtml = "<p style='color:rgba(30,33,36,0.55);text-align:center;padding:16px;'>Authentication methods data was not available. Ensure <code>Policy.Read.All</code> is granted.</p>"
}

# Guest stats
$gaGuestCount = 0; $gaTotalUsers = 0; $gaInvitePolicy = 'Unknown'
if ($data.PSObject.Properties.Name -contains 'GuestAccess' -and $data.GuestAccess) {
    $ga = $data.GuestAccess
    try { $gaGuestCount = [int]$ga.GuestUserCount } catch {}
    try { $gaTotalUsers = [int]$ga.TotalUsers }     catch {}
    $gaInvitePolicy = [string]$ga.GuestInviteSettings
}
$gaGuestPct = if ($gaTotalUsers -gt 0) { [math]::Round(100.0 * $gaGuestCount / $gaTotalUsers, 1) } else { 0 }
# Friendly name for invite policy
$invitePolicyFriendly = switch ($gaInvitePolicy) {
    'adminsAndGuestInviters'     { 'Admins and guest inviters only' }
    'adminsGuestInvitersAndAllMembers' { 'All members can invite' }
    'everyone'                    { 'Everyone (including guests)' }
    'none'                        { 'No one can invite' }
    default                       { $gaInvitePolicy }
}

[void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='page-h1'>Authentication Methods</h1>
  <p class='body-text'>The cards below show which authentication methods are enabled or disabled at the tenant level. Methods such as SMS and Voice are considered weaker and should be disabled in favour of phishing-resistant alternatives like FIDO2 or Microsoft Authenticator.</p>
  <div class='am-grid'>
    $amCardsHtml
  </div>
  <h1 class='page-h1' style='margin-top:28px;'>External & Guest Access</h1>
  <p class='body-text'>Guest users are external identities granted access to tenant resources. The metrics below summarise the current guest footprint and the tenant invitation policy.</p>
  <div class='ga-grid'>
    <div class='ga-card'>
      <div class='ga-num'>$gaGuestCount</div>
      <div class='ga-label'>Guest Users</div>
    </div>
    <div class='ga-card'>
      <div class='ga-num'>$gaGuestPct%</div>
      <div class='ga-label'>of Total Directory</div>
    </div>
    <div class='ga-card'>
      <div class='ga-num' style='font-size:16px;'>$(Enc $invitePolicyFriendly)</div>
      <div class='ga-label'>Invitation Policy</div>
    </div>
  </div>
</section>
"@)

# ===========================================================================
# Page 7: Microsoft Security Baseline (radar / spider chart)
# ===========================================================================
$radarAxes = @()
foreach ($sec in $SectionOrder) {
    $secCtrls = @($controls | Where-Object { [string]$_.Section -eq $sec })
    $passCount = @($secCtrls | Where-Object { ("$($_.Status)").Trim().ToUpper() -eq 'PASS' }).Count
    $radarAxes += @{ Label = ($SectionTitles[$sec] -replace '^\d+\s*-\s*',''); Value = $passCount }
}
$radarMax = ($radarAxes | ForEach-Object { [int]$_.Value } | Measure-Object -Maximum).Maximum
if ($radarMax -lt 5) { $radarMax = 5 }
$radarSvg = Get-RadarSvg -Axes $radarAxes -Max $radarMax

[void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='page-h1'>Microsoft Security Baseline</h1>
  <p class='body-text'>The chart below plots the number of security controls currently passing across each Microsoft 365 workload. A larger, more balanced shape indicates broader coverage of the recommended security baseline. Areas where the shape contracts toward the center highlight workloads that would benefit from additional hardening.</p>
  <div class='radar-wrap'>
    $radarSvg
    <div class='radar-legend'><span class='radar-swatch'></span> Controls passing (current state)</div>
  </div>
</section>
"@)

# ===========================================================================
# Page 8: Microsoft Security Baseline Overview (tier gauges)
# ===========================================================================
# Static CheckId -> maturity tier classification
$TierMap = @{
    'Essentials' = @('1.1','1.2','1.3','1.4','1.5','1.6','1.9','1.12','2.1','2.2','2.3','2.4','2.5','2.6','3.1','4.2','4.5','4.6','4.7','5.1','6.2','6.3','6.4','6.5','6.6','7.1','7.2')
    'Core'       = @('1.7','1.8','1.10','1.11','1.17','1.18','3.2','3.3','3.4','3.5','4.1','4.3','4.8','6.7','7.3')
    'Premium'    = @('1.13','1.19','1.20','3.6','3.7','5.2','6.8')
    'Advanced'   = @('1.14','1.15','1.16','1.21','1.22','1.23','1.24','4.4','4.9','4.10','6.1','6.9','7.4','7.5')
}
$TierOrder = @('Essentials','Core','Premium','Advanced')
$tierCardsHtml = New-Object System.Text.StringBuilder
foreach ($tier in $TierOrder) {
    $ids = $TierMap[$tier]
    $passed = 0; $failed = 0; $assumed = 0; $notset = 0
    foreach ($id in $ids) {
        if (-not $parentStatusMap.ContainsKey($id)) { continue }
        switch ($parentStatusMap[$id]) {
            'PASS'   { $passed++ }
            'FAIL'   { $failed++ }
            'WARN'   { $failed++ }
            'MANUAL' { $notset++ }
            'NOTSET' { $notset++ }
            default  { $notset++ }
        }
    }
    $total = $passed + $failed + $assumed + $notset
    $gauge = Get-HalfGaugeSvg -Value $passed -Total $total -Label $tier
    [void]$tierCardsHtml.Append(@"
<div class='tier-card'>
  $gauge
  <div class='tier-legend'>
    <div class='tl-row'><span class='tl-ico ok'></span>$passed Passed</div>
    <div class='tl-row'><span class='tl-ico bad'></span>$failed Failed</div>
    <div class='tl-row'><span class='tl-ico amber'></span>$assumed Assumed Risk</div>
    <div class='tl-row'><span class='tl-ico gray'></span>$notset Not Set</div>
  </div>
</div>
"@)
}

[void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='page-h1'>Security Baseline Overview</h1>
  <p class='body-text'>Security controls are grouped into four maturity tiers. Each gauge shows how many controls in that tier are currently passing out of the controls that could be evaluated. "Not Set" indicates controls that require manual review or were not configured.</p>
  <div class='tier-grid'>
    $($tierCardsHtml.ToString())
  </div>
  <p class='mf-note'>Controls requiring manual verification are reported as "Not Set" until reviewed.</p>
</section>
"@)

# ---- Page 9: Prioritised Recommendations ----------------------------------
$sevRank = @{ 'critical' = 0; 'high' = 1; 'medium' = 2; 'low' = 3 }
function Get-WorstSeverity($items) {
    $worst = 'low'; $worstRank = 99
    foreach ($it in $items) {
        $s = ("$($it.Severity)").Trim().ToLower()
        $r = if ($sevRank.ContainsKey($s)) { $sevRank[$s] } else { 4 }
        if ($r -lt $worstRank) { $worstRank = $r; $worst = $s }
    }
    return $worst
}
$recGroups = @()
foreach ($grp in ($failedControls | Group-Object CheckId)) {
    $items = @($grp.Group)
    $title = $ParentTitles[[string]$grp.Name]
    if ([string]::IsNullOrWhiteSpace($title)) { $title = ($items | Select-Object -First 1).Description }
    $detail = (($items | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Detail) } | Select-Object -First 1)).Detail
    $recGroups += [pscustomobject]@{
        CheckId  = [string]$grp.Name
        Title    = $title
        Detail   = [string]$detail
        Severity = Get-WorstSeverity -items $items
    }
}
$topRecs = $recGroups | Sort-Object { if ($sevRank.ContainsKey($_.Severity)) { $sevRank[$_.Severity] } else { 4 } }, CheckId | Select-Object -First 8

if ($topRecs.Count -gt 0) {
    $recNum = 0
    $recRowsHtml = ($topRecs | ForEach-Object {
        $recNum++
        $badge = Get-SeverityBadge -Severity $_.Severity
        $detailHtml = if (-not [string]::IsNullOrWhiteSpace($_.Detail)) { "<div class='rec-detail'>$(Enc $_.Detail)</div>" } else { '' }
        "<div class='rec-item'><div class='rec-num'>$recNum</div><div class='rec-body'><div class='rec-title'>$badge $(Enc $_.CheckId) &mdash; $(Enc $_.Title)</div>$detailHtml</div></div>"
    }) -join "`n"
    $recIntro = "The items below are drawn from the $totalFailCount failing control(s) found across the full assessment, ordered by severity. Addressing these first gives the fastest reduction in real-world risk."
} else {
    $recRowsHtml = "<div class='rec-empty'>No failing controls were found across the areas we assessed &mdash; a strong result. Review the detailed findings in the following pages for any items marked for manual verification.</div>"
    $recIntro = "No failing controls were found across the areas we assessed."
}

[void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='page-h1'>Recommended Next Steps</h1>
  <p class='body-text'>$recIntro</p>
  <div class='rec-list'>
    $recRowsHtml
  </div>
</section>
"@)

# ---- Section check pages -------------------------------------------------
foreach ($sec in $SectionOrder) {
    $secControls = @($controls | Where-Object { [string]$_.Section -eq $sec })
    if ($secControls.Count -eq 0) { continue }

    $secTitle = $SectionTitles[$sec]
    if ([string]::IsNullOrWhiteSpace($secTitle)) { $secTitle = $sec }

    # group by CheckId preserving numeric order
    $groups = $secControls | Group-Object CheckId
    $groups = $groups | Sort-Object `
        @{ Expression = { $p = ($_.Name -split '\.'); if ($p.Count -ge 1 -and $p[0] -match '^\d+$') { [int]$p[0] } else { 999 } } }, `
        @{ Expression = { $p = ($_.Name -split '\.'); if ($p.Count -ge 2 -and $p[1] -match '^\d+$') { [int]$p[1] } else { 999 } } }

    # Render each check-group's HTML once, keyed by group, before deciding
    # how to paginate - pagination (below) is now explicit rather than left
    # to the browser.
    $groupHtmlByName = [ordered]@{}
    $groupUnitsByName = @{}
    foreach ($g in $groups) {
        $items = @($g.Group)
        $statuses = $items | ForEach-Object { ("$($_.Status)").Trim().ToUpper() }
        # parent status: FAIL if any FAIL, else MANUAL if all MANUAL/NOTSET, else PASS if all PASS, else WARN
        $hasFail = ($statuses -contains 'FAIL')
        $allManual = ($statuses.Count -gt 0) -and -not ($statuses | Where-Object { $_ -notmatch '^(MANUAL|NOTSET|NOT SET|NA|N/A)$' })
        $allPass = ($statuses.Count -gt 0) -and -not ($statuses | Where-Object { $_ -ne 'PASS' })
        if ($hasFail) { $parentStatus = 'FAIL' }
        elseif ($allPass) { $parentStatus = 'PASS' }
        elseif ($allManual) { $parentStatus = 'MANUAL' }
        else { $parentStatus = 'WARN' }

        $parentTitle = $ParentTitles[[string]$g.Name]
        if ([string]::IsNullOrWhiteSpace($parentTitle)) {
            $parentTitle = ($items | Select-Object -First 1).Description
        }
        $pIcon = Get-StatusIcon -Status $parentStatus -Size 34

        $subs = New-Object System.Text.StringBuilder
        foreach ($it in $items) {
            $st = ("$($it.Status)").Trim().ToUpper()
            $sIcon = Get-StatusIcon -Status $st -Size 22
            $badge = ''
            if ($st -eq 'FAIL' -or $st -like 'WARN*') { $badge = Get-SeverityBadge -Severity $it.Severity }
            $detail = [string]$it.Detail
            $detailHtml = ''
            if (-not [string]::IsNullOrWhiteSpace($badge) -or -not [string]::IsNullOrWhiteSpace($detail)) {
                $detailHtml = "<div class='sub-detail'>$badge<span class='sub-detail-text'>$(Enc $detail)</span></div>"
            }
            [void]$subs.Append("<div class='subcheck'><div class='sub-icon'>$sIcon</div><div class='sub-body'><div class='sub-title'>$(Enc $it.Description)</div>$detailHtml</div></div>")
        }

        $groupHtmlByName[[string]$g.Name] = "<div class='check-group'><div class='check-head'><div class='check-icon'>$pIcon</div><div class='check-title'>$(Enc $g.Name) - $(Enc $parentTitle)</div></div><div class='check-subs'>$($subs.ToString())</div></div>"
        # Rough per-group "weight" for pagination below: one unit for the
        # group's own header row, plus one per sub-check row. Not pixel-
        # exact (detail text length varies row height further), but good
        # enough to keep pages comfortably under one physical page's worth
        # of content - a little unused white space at the bottom of a page
        # is a far smaller problem than the one this is fixing.
        $groupUnitsByName[[string]$g.Name] = 1 + $items.Count
    }

    # If every control in this section was skipped for the same
    # tier-limitation reason (Exchange/Teams under the Graph-only Essential
    # tier - see M365-SecurityAssessment.ps1's SkipExchange/$HasTeamsModule),
    # say so once, prominently, right under the section heading - rather
    # than only leaving the customer to notice the same explanation
    # repeated in every individual check's fine print below.
    # The marker can land in either field depending on which check wrote it -
    # M365-SecurityAssessment.ps1's Exchange checks put it in Description
    # (the parent check title), Teams checks put it in Detail. Check both.
    $tierGateMarker = 'Not included in this assessment tier'
    $allTierGated = ($secControls.Count -gt 0) -and
        -not (@($secControls | Where-Object {
            $_.Detail -notlike "*$tierGateMarker*" -and $_.Description -notlike "*$tierGateMarker*"
        })).Count
    $tierNoteHtml = ''
    if ($allTierGated) {
        $tierNoteHtml = @"
  <div class='scope-note tier-note'><strong>Not included in this assessment tier:</strong> the checks below require certificate-based access, included in our Advanced assessment - see the individual items for detail, or reply to your report email to ask about upgrading.</div>
"@
    }

    # Explicit pagination: pack groups into page-sized chunks by running
    # unit total, rather than emitting one <section> per section and
    # letting the browser fragment it across physical pages if it
    # overflows. Confirmed live (twice) that a fragmented continuation
    # page does not reliably get .page's own padding reapplied in this
    # Edge print-to-pdf pipeline, regardless of @page CSS margin settings -
    # continuation pages kept rendering flush against the page edge. Every
    # chunk below becomes its own genuine top-level <section class='page'>,
    # getting real padding the same way the cover and closing pages always
    # correctly have, since neither of those has ever depended on browser
    # fragmentation.
    $maxUnitsPerPage = 11
    $pageChunks = New-Object System.Collections.Generic.List[object]
    $currentChunk = New-Object System.Collections.Generic.List[object]
    $currentUnits = 0
    foreach ($g in $groups) {
        $gName = [string]$g.Name
        $groupUnits = $groupUnitsByName[$gName]
        if ($currentChunk.Count -gt 0 -and ($currentUnits + $groupUnits) -gt $maxUnitsPerPage) {
            $pageChunks.Add($currentChunk)
            $currentChunk = New-Object System.Collections.Generic.List[object]
            $currentUnits = 0
        }
        [void]$currentChunk.Add($gName)
        $currentUnits += $groupUnits
    }
    if ($currentChunk.Count -gt 0) { $pageChunks.Add($currentChunk) }

    $chunkIndex = 0
    foreach ($chunk in $pageChunks) {
        $chunkIndex++
        $chunkGroupsHtml = ($chunk | ForEach-Object { $groupHtmlByName[$_] }) -join "`n"
        $chunkTitle = if ($chunkIndex -eq 1) { Enc $secTitle } else { "$(Enc $secTitle) (continued)" }
        # The tier-gate banner is a section-level summary - show it only
        # once, on the section's first page.
        $chunkTierNoteHtml = if ($chunkIndex -eq 1) { $tierNoteHtml } else { '' }
        [void]$pages.Append(@"
<section class='page section-page'>
  $Watermark
  <h1 class='section-h1'>$chunkTitle</h1>
$chunkTierNoteHtml  <div class='check-list'>
    $chunkGroupsHtml
  </div>
</section>
"@)
    }
}

# ---- Closing page -----------------------------------------------------------
if ($IsAdvancedTier) {
    $ctaPanelTitle = 'Next steps on these findings'
    $ctaPanelText  = 'As part of this engagement, we can prioritise and implement the recommendations in this report, or scope further work covering areas outside what this assessment can reach (see Scope &amp; Methodology).'
    $accessText    = "Our application held read-only Microsoft Graph, Exchange Online and Microsoft Teams permissions for the duration of this assessment - broader than our free assessment, to reach the additional depth in this report. You're welcome to remove it at any time from Enterprise Applications in the Azure portal - this report has already been generated, so removing access does not affect it."
} else {
    $ctaPanelTitle = 'Want help closing these gaps?'
    $ctaPanelText  = 'We can prioritise and implement the recommendations in this report, or scope a deeper assessment covering areas outside what a free, read-only review can reach (see Scope &amp; Methodology). No obligation - just a conversation.'
    $accessText    = "Our application held only read-only Microsoft Graph permissions for the duration of this assessment. You're welcome to remove it at any time from Enterprise Applications in the Azure portal - this report has already been generated, so removing access does not affect it."
}
[void]$pages.Append(@"
<section class='page'>
  $Watermark
  <h1 class='cta-title'>Where to go from here</h1>
  <p class='cta-body'>Thank you for trusting Camelot Digital Group with this assessment. The findings in this report are yours to keep and act on however suits you best - with your own team, another provider, or with us.</p>
  <div class='cta-panel'>
    <div class='cta-panel-title'>$ctaPanelTitle</div>
    <p class='cta-panel-text'>$ctaPanelText</p>
    <div class='cta-contact'>Reply to the email that delivered this report, or reach us at <a href='mailto:hello@camelotdigitalgroup.com'>hello@camelotdigitalgroup.com</a></div>
  </div>
  <div class='cta-panel'>
    <div class='cta-panel-title'>Removing our access</div>
    <p class='cta-panel-text'>$accessText</p>
  </div>
</section>
"@)

# ===========================================================================
# CSS
# ===========================================================================
$css = @'
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=EB+Garamond:ital,wght@0,400;0,500;0,600;0,700;1,400&display=swap');
* { margin:0; padding:0; box-sizing:border-box; print-color-adjust:exact; -webkit-print-color-adjust:exact; }
html, body { background:#F6F3EC; font-family:'Space Grotesk', 'Segoe UI', Arial, Helvetica, sans-serif; color:#1E2124; }
.page {
  /* min-height must account for @page's top margin below (48px, matching
     .page's own top padding): the printable area on every physical page
     is 297mm minus that margin, not the full 297mm. Leaving this at a
     flat 297mm meant EVERY .page div that filled a full page overflowed
     onto a second, near-blank physical page - not just genuinely
     multi-page sections. Confirmed live: ~22 logical page sections
     rendered as ~45 physical PDF pages. */
  width:210mm; min-height:calc(297mm - 48px); background:#F6F3EC;
  background-image: radial-gradient(rgba(212, 169, 79, 0.09) 1px, transparent 1.5px);
  background-size: 22px 22px;
  padding:48px 56px; page-break-after:always; position:relative; overflow:hidden;
}
.page:last-child { page-break-after:auto; }

/* Watermark - the real Camelot crest mark, centered, low opacity, behind
   all content on every page. Matches website/globals.css's own
   .heritage-pattern texture applied to .page above - both are genuine
   brand elements, not invented decoration. */
.page-watermark { position:absolute !important; left:50%; top:50%; width:380px; height:380px; transform:translate(-50%,-50%); opacity:0.05; z-index:0; }
.page-watermark svg { width:100%; height:100%; }
.page > * { position:relative; z-index:1; }

/* Cover */
.cover { display:flex; flex-direction:column; }
.logo { display:flex; align-items:center; }
.logo svg { width:300px; height:auto; }
.cover-top { margin-bottom:120px; }
.cover-body { margin-top:60px; }
.cover-date { font-size:20px; color:#B8902E; margin-bottom:40px; letter-spacing:1px; font-weight:600; }
.cover-name { font-family:'EB Garamond', Georgia, serif; font-size:70px; font-weight:700; line-height:1.02; color:#071A33; }
.cover-title { font-family:'EB Garamond', Georgia, serif; font-size:56px; font-weight:400; font-style:italic; line-height:1.05; color:#071A33; margin-top:2px; }
.cover-tier-badge { display:inline-block; margin-top:20px; padding:6px 16px; background:#071A33; color:#D4A94F; font-size:13px; font-weight:600; letter-spacing:1.5px; text-transform:uppercase; border-radius:20px; }
.cover-rule { height:2px; background:linear-gradient(to right,#D4A94F,rgba(212,169,79,0)); width:92%; margin:34px 0 26px; }
.cover-exec { font-size:24px; color:rgba(30,33,36,0.68); font-weight:300; }
.cover-foot { margin-top:auto; color:rgba(30,33,36,0.55); font-size:12px; letter-spacing:1px; }

/* Headings */
.page-h1 { font-family:'EB Garamond', Georgia, serif; font-size:46px; font-weight:600; color:#071A33; margin-bottom:34px; }
.section-h1 { font-family:'EB Garamond', Georgia, serif; font-size:26px; font-weight:700; color:#071A33; margin-bottom:28px; }
.body-text { font-size:15px; line-height:1.6; color:#1E2124; margin:20px 0; max-width:95%; }

/* Secure score */
.ss-row { display:flex; align-items:center; gap:30px; margin:30px 0 10px; }
.ss-gauge { flex:0 0 auto; }
.ss-chart { flex:1 1 auto; }
.cmp-wrap { display:flex; align-items:center; gap:18px; margin-top:60px; }
.cmp-label-left, .cmp-label-right { font-weight:700; font-size:14px; color:#071A33; text-align:center; white-space:nowrap; }
.cmp-track-wrap { position:relative; flex:1 1 auto; padding:34px 0; }
.cmp-track { height:16px; border-radius:8px; background:linear-gradient(to right,#dc2626,#ef4444,#d97706,#eab308,#65a30d,#16a34a); position:relative; }
.cmp-dot { position:absolute; top:50%; width:20px; height:20px; border-radius:50%; background:#071A33; border:2px solid #F6F3EC; transform:translate(-50%,-50%); box-shadow:0 0 0 1px #B8902E; }
.cmp-marker-top, .cmp-marker-bot { position:absolute; transform:translateX(-50%); }
.cmp-marker-top { top:0; } .cmp-marker-bot { bottom:0; }
.cmp-tip { display:inline-block; background:#071A33; color:#F6F3EC; font-size:12px; font-weight:600; padding:4px 9px; border-radius:5px; white-space:nowrap; }

/* Metric rows (user/device health) */
.metric-list { display:flex; flex-direction:column; gap:34px; margin-top:10px; }
.metric-row { display:grid; grid-template-columns:110px 260px 1fr; align-items:center; gap:8px; break-inside:avoid; page-break-inside:avoid; }
.metric-circle { width:88px; height:88px; border-radius:50%; background:#FFFFFF; border:1px solid rgba(7,26,51,0.14); display:flex; align-items:center; justify-content:center; }
.metric-num { font-size:44px; font-weight:600; color:#071A33; display:flex; align-items:center; gap:8px; line-height:1; }
.metric-ind { display:inline-flex; }
.metric-label { font-size:15px; color:#1E2124; margin-top:8px; line-height:1.3; font-weight:600; }
.metric-desc { font-size:14px; color:rgba(30,33,36,0.68); line-height:1.5; }

/* Applications & Data */
.ad-sub { font-size:20px; font-weight:700; color:#071A33; margin:28px 0 14px; font-family:'EB Garamond', Georgia, serif; }
.share-row { display:flex; align-items:flex-start; gap:16px; break-inside:avoid; page-break-inside:avoid; }
.share-icon { flex:0 0 auto; margin-top:2px; }
.share-title { font-size:16px; font-weight:600; color:#071A33; }
.share-desc { font-size:14px; color:rgba(30,33,36,0.68); margin-top:3px; max-width:90%; }
.sp-list { display:flex; flex-direction:column; gap:16px; margin-top:6px; }
.sp-row { display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid rgba(7,26,51,0.14); padding-bottom:10px; break-inside:avoid; page-break-inside:avoid; }
.sp-name { font-size:15px; font-weight:700; color:#071A33; }
.sp-count-num { font-size:18px; color:#071A33; }
.sp-count-lbl { font-size:12px; color:rgba(30,33,36,0.55); }
.sp-empty { color:rgba(30,33,36,0.55); font-size:14px; }

/* SharePoint site inventory table */
.spi-page { font-size:14px; color:rgba(30,33,36,0.55); font-weight:400; }
.spi-table { width:100%; border-collapse:collapse; margin-top:14px; font-size:12.5px; }
.spi-table thead th { text-align:left; color:#B8902E; font-weight:700; border-bottom:1px solid rgba(7,26,51,0.18); padding:8px 10px; font-size:12px; text-transform:uppercase; letter-spacing:.4px; }
.spi-table tbody tr { break-inside:avoid; page-break-inside:avoid; }
.spi-table tbody td { padding:7px 10px; border-bottom:1px solid rgba(7,26,51,0.10); color:#1E2124; vertical-align:top; }
.spi-table tbody tr:nth-child(even) td { background:rgba(7,26,51,0.035); }
.spi-name { font-weight:600; color:#071A33; width:38%; word-break:break-word; }
.spi-url { color:rgba(30,33,36,0.55); word-break:break-all; }

/* Section check pages */
.check-list { display:flex; flex-direction:column; gap:22px; }
.check-group { break-inside:avoid; page-break-inside:avoid; }
.check-head { display:flex; align-items:flex-start; gap:14px; }
.check-icon { flex:0 0 auto; margin-top:1px; }
.check-title { font-size:15px; font-weight:700; color:#071A33; line-height:1.35; padding-top:6px; }
.check-subs { margin:10px 0 0 54px; display:flex; flex-direction:column; gap:14px; }
.subcheck { display:flex; align-items:flex-start; gap:12px; break-inside:avoid; page-break-inside:avoid; }
.sub-icon { flex:0 0 auto; margin-top:1px; }
.sub-title { font-size:14px; font-weight:600; color:#071A33; line-height:1.3; }
.sub-detail { margin-top:4px; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
.sub-detail-text { font-size:12px; color:rgba(30,33,36,0.68); line-height:1.4; }
.badge { font-size:10px; padding:1px 6px; border-radius:3px; font-weight:600; display:inline-block; line-height:1.5; }

/* Email Health */
.eh-table { width:100%; border-collapse:collapse; margin:14px 0 10px; }
.eh-table thead th { font-size:12px; font-weight:700; color:#B8902E; text-transform:uppercase; letter-spacing:1px; text-align:center; padding:8px 6px; border-bottom:2px solid rgba(7,26,51,0.18); }
.eh-table thead th.eh-dom { text-align:left; }
.eh-table tbody tr { break-inside:avoid; page-break-inside:avoid; }
.eh-table td { padding:7px 6px; border-bottom:1px solid rgba(7,26,51,0.10); font-size:13px; vertical-align:middle; }
.eh-table td.eh-dom { text-align:left; color:#071A33; font-weight:600; }
.eh-table td.eh-c { text-align:center; }
.eh-table td.eh-c svg { vertical-align:middle; }
.eh-default { color:#16A34A; font-weight:700; }
.eh-no { color:rgba(30,33,36,0.55); }
.badge-pass { display:inline-block; padding:2px 8px; border-radius:4px; font-size:10px; font-weight:700; letter-spacing:1px; background:#E8F1EA; color:#166534; }
.badge-fail { display:inline-block; padding:2px 8px; border-radius:4px; font-size:10px; font-weight:700; letter-spacing:1px; background:#DC2626; color:#ffffff; }
.badge-na   { display:inline-block; padding:2px 8px; border-radius:4px; font-size:10px; font-weight:700; letter-spacing:1px; background:rgba(7,26,51,0.10); color:rgba(30,33,36,0.55); }

/* Licensing table */
.lic-table { width:100%; border-collapse:collapse; margin:14px 0 10px; }
.lic-table thead th { font-size:11px; font-weight:700; color:#B8902E; text-transform:uppercase; letter-spacing:1px; text-align:left; padding:8px 6px; border-bottom:2px solid rgba(7,26,51,0.18); }
.lic-table tbody tr { break-inside:avoid; page-break-inside:avoid; }
.lic-table td { padding:7px 6px; border-bottom:1px solid rgba(7,26,51,0.10); font-size:12px; vertical-align:middle; color:#071A33; }
.lic-table td.lic-name { font-weight:600; max-width:240px; overflow:hidden; text-overflow:ellipsis; }
.lic-bar-track { width:120px; height:8px; background:rgba(7,26,51,0.10); border-radius:4px; display:inline-block; vertical-align:middle; margin-right:8px; }
.lic-bar-fill { height:100%; border-radius:4px; }
.lic-pct { font-size:11px; color:rgba(30,33,36,0.55); }

/* Auth methods */
.am-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:8px; margin:14px 0; }
.am-card { background:#FFFFFF; border:1px solid rgba(212,169,79,0.22); box-shadow:0 1px 2px rgba(0,0,0,0.04), 0 10px 24px -16px rgba(7,26,51,0.28); border-radius:8px; padding:12px 10px; text-align:center; break-inside:avoid; page-break-inside:avoid; }
.am-name { font-size:11px; color:#1E2124; font-weight:600; text-transform:capitalize; }
.am-state { font-size:10px; font-weight:700; margin-top:4px; text-transform:uppercase; letter-spacing:1px; }
.am-enabled  { color:#16A34A; }
.am-disabled { color:#DC2626; }
.am-other    { color:rgba(30,33,36,0.55); }

/* Guest access */
.ga-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:12px; margin:14px 0; }
.ga-card { background:#FFFFFF; border:1px solid rgba(212,169,79,0.22); box-shadow:0 1px 2px rgba(0,0,0,0.04), 0 10px 24px -16px rgba(7,26,51,0.28); border-radius:8px; padding:16px; text-align:center; break-inside:avoid; page-break-inside:avoid; }
.ga-num { font-size:32px; font-weight:800; color:#B8902E; line-height:1; }
.ga-label { font-size:11px; color:rgba(30,33,36,0.55); margin-top:6px; }
.eh-pol { display:inline-block; margin-left:6px; font-size:11px; color:rgba(30,33,36,0.55); vertical-align:middle; }
.eh-empty { text-align:center; color:rgba(30,33,36,0.55); padding:20px; }
.mf-bars { display:flex; flex-direction:column; gap:16px; margin-top:8px; }
.mf-item { break-inside:avoid; page-break-inside:avoid; }
.mf-top { display:flex; justify-content:space-between; align-items:baseline; margin-bottom:6px; }
.mf-label { font-size:14px; color:#1E2124; }
.mf-num { font-size:16px; font-weight:700; color:#071A33; }
.mf-track { height:14px; background:rgba(7,26,51,0.10); border-radius:7px; overflow:hidden; }
.mf-fill { height:100%; border-radius:7px; }
.mf-note { font-size:12px; color:rgba(30,33,36,0.55); margin-top:14px; }

/* Radar */
.radar-wrap { display:flex; flex-direction:column; align-items:center; margin-top:6px; }
.radar-legend { margin-top:10px; font-size:13px; color:#1E2124; display:flex; align-items:center; gap:8px; }
.radar-swatch { display:inline-block; width:22px; height:12px; background:#D4A94F; opacity:0.5; border:1px solid #D4A94F; border-radius:2px; }

/* Baseline overview tiers */
.tier-grid { display:grid; grid-template-columns:1fr 1fr; gap:36px 28px; margin-top:20px; }
.tier-card { display:flex; align-items:center; gap:20px; background:#FFFFFF; border:1px solid rgba(212,169,79,0.22); box-shadow:0 1px 2px rgba(0,0,0,0.04), 0 10px 24px -16px rgba(7,26,51,0.28); border-radius:12px; padding:18px 22px; break-inside:avoid; page-break-inside:avoid; }
.tier-legend { display:flex; flex-direction:column; gap:7px; }
.tl-row { display:flex; align-items:center; gap:9px; font-size:14px; color:#1E2124; }
.tl-ico { display:inline-block; width:14px; height:14px; border-radius:50%; flex:0 0 auto; }
.tl-ico.ok { background:#16A34A; }
.tl-ico.bad { background:#DC2626; }
.tl-ico.amber { background:#D97706; }
.tl-ico.gray { background:rgba(30,33,36,0.35); }

/* Executive summary */
.exec-summary { font-size:17px; line-height:1.7; color:#1E2124; margin:22px 0 0; max-width:92%; }
.exec-summary strong { color:#071A33; }
.exec-stat-row { display:grid; grid-template-columns:repeat(4,1fr); gap:16px; margin-top:36px; }
.exec-stat { background:#FFFFFF; border:1px solid rgba(212,169,79,0.22); box-shadow:0 1px 2px rgba(0,0,0,0.04), 0 10px 24px -16px rgba(7,26,51,0.28); border-radius:10px; padding:18px 16px; text-align:center; }
.exec-stat-num {
  font-family:'EB Garamond', Georgia, serif; font-size:36px; font-weight:700; line-height:1;
  background:linear-gradient(135deg, #E7C878 0%, #D4A94F 48%, #B8902E 100%);
  -webkit-background-clip:text; background-clip:text; -webkit-text-fill-color:transparent; color:#B8902E;
}
.exec-stat-label { font-size:11px; color:rgba(30,33,36,0.55); margin-top:8px; text-transform:uppercase; letter-spacing:.6px; }

/* Scope & methodology */
.scope-list { display:flex; flex-direction:column; gap:16px; margin-top:24px; }
.scope-item { display:flex; align-items:flex-start; gap:14px; break-inside:avoid; page-break-inside:avoid; }
.scope-mark { flex:0 0 auto; width:8px; height:8px; border-radius:50%; background:#D4A94F; margin-top:8px; }
.scope-text { font-size:14.5px; line-height:1.6; color:#1E2124; }
.scope-text strong { color:#071A33; }
.scope-note { margin-top:30px; padding:18px 20px; background:#FFFFFF; border:1px solid rgba(212,169,79,0.22); box-shadow:0 1px 2px rgba(0,0,0,0.04), 0 10px 24px -16px rgba(7,26,51,0.28); border-left:4px solid #D4A94F; border-radius:6px; font-size:13.5px; line-height:1.6; color:rgba(30,33,36,0.68); }
.scope-note.tier-note { margin-top:0; margin-bottom:22px; break-inside:avoid; page-break-inside:avoid; }

/* Recommendations */
.rec-list { display:flex; flex-direction:column; gap:18px; margin-top:26px; }
.rec-item { display:flex; align-items:flex-start; gap:16px; background:#FFFFFF; border:1px solid rgba(212,169,79,0.22); box-shadow:0 1px 2px rgba(0,0,0,0.04), 0 10px 24px -16px rgba(7,26,51,0.28); border-radius:10px; padding:16px 18px; break-inside:avoid; page-break-inside:avoid; }
.rec-num { flex:0 0 auto; width:32px; height:32px; border-radius:50%; background:#071A33; color:#D4A94F; font-family:'EB Garamond', Georgia, serif; font-weight:700; font-size:16px; display:flex; align-items:center; justify-content:center; }
.rec-body { flex:1 1 auto; }
.rec-title { font-size:14.5px; font-weight:600; color:#071A33; line-height:1.35; }
.rec-detail { font-size:13px; color:rgba(30,33,36,0.68); margin-top:4px; line-height:1.5; }
.rec-empty { text-align:center; color:rgba(30,33,36,0.68); padding:40px 20px; font-size:15px; }

/* Closing / CTA */
.cta-title { font-family:'EB Garamond', Georgia, serif; font-size:38px; font-weight:600; color:#071A33; margin-bottom:18px; }
.cta-body { font-size:15.5px; line-height:1.7; color:#1E2124; max-width:80%; }
.cta-panel { margin-top:40px; background:#FFFFFF; border:1px solid rgba(212,169,79,0.22); box-shadow:0 1px 2px rgba(0,0,0,0.04), 0 10px 24px -16px rgba(7,26,51,0.28); border-radius:12px; padding:28px 32px; }
.cta-panel-title { font-size:18px; font-weight:700; color:#071A33; margin-bottom:10px; }
.cta-panel-text { font-size:14px; line-height:1.6; color:rgba(30,33,36,0.68); }
.cta-contact { margin-top:22px; font-size:15px; color:#071A33; font-weight:600; }
.cta-contact a { color:#B8902E; text-decoration:none; }
'@

# ===========================================================================
# Assemble HTML
# ===========================================================================
$html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<meta name='viewport' content='width=device-width, initial-scale=1.0'>
<title>$(Enc $CustomerName) - Cloud Assessment Report</title>
<style>
/* Top-only @page margin, matching .page's own 48px top padding exactly.
   .page's padding covers spacing for every page that IS its own
   <section> element (cover, closing page, first page of each section).
   It's NOT reapplied when a section's content naturally overflows one
   .page div onto a second physical page (a DOM-fragmentation limit of
   CSS paged media, not something .page's own padding can fix - that
   padding only renders once, at the true top of the div's box, not at
   each page it visually spans) - confirmed live: those continuation
   pages start flush against the physical page edge. This adds the
   missing breathing room back for exactly that case. First tried a
   smaller 24px here (half of .page's 48px) - confirmed live that
   continuation pages still read as visibly tighter than a genuine
   page start right next to them, so this now matches exactly rather
   than approximating. .page's min-height below is adjusted to match
   (calc(297mm - 48px)) so pages that fill a full page don't overflow
   onto a spurious near-blank one. */
@page { size:A4 portrait; margin:48px 0 0 0; }
$css
</style>
</head>
<body>
$($pages.ToString())
</body>
</html>
"@

# ---------------------------------------------------------------------------
# Write files
# ---------------------------------------------------------------------------
$stamp = $reportDate.ToString('yyyyMMdd_HHmmss')
$htmlPath = Join-Path $OutputFolder "Assessment_Report_$stamp.html"
$pdfPath  = Join-Path $OutputFolder "Assessment_Report_$stamp.pdf"

$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($htmlPath, $html, $utf8)
Write-Host "HTML report written to: $htmlPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Convert to PDF via Edge / Chrome headless
# ---------------------------------------------------------------------------
$pf   = ${env:ProgramFiles}
$pfx86 = ${env:ProgramFiles(x86)}
$browserCandidates = @()
if ($pf)    { $browserCandidates += (Join-Path $pf    'Microsoft\Edge\Application\msedge.exe') }
if ($pfx86) { $browserCandidates += (Join-Path $pfx86 'Microsoft\Edge\Application\msedge.exe') }
if ($pf)    { $browserCandidates += (Join-Path $pf    'Google\Chrome\Application\chrome.exe') }
if ($pfx86) { $browserCandidates += (Join-Path $pfx86 'Google\Chrome\Application\chrome.exe') }
# Cross-platform fallbacks (also lets the script run/test on non-Windows hosts)
$browserCandidates += @('/usr/bin/google-chrome','/usr/bin/chromium','/usr/local/bin/chromium','/usr/bin/microsoft-edge')
$browserExe = $browserCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1

if ($browserExe) {
    Write-Host "Converting to PDF using: $browserExe" -ForegroundColor Cyan
    $htmlFull = (Resolve-Path -LiteralPath $htmlPath).ProviderPath
    $fileUri = (New-Object System.Uri($htmlFull)).AbsoluteUri
    $procArgs = @(
        '--headless=new'
        '--disable-gpu'
        '--no-sandbox'
        '--no-pdf-header-footer'
        "--print-to-pdf=$pdfPath"
        $fileUri
    )
    try {
        $proc = Start-Process -FilePath $browserExe -ArgumentList $procArgs -Wait -PassThru
        Start-Sleep -Milliseconds 500
        if (Test-Path -LiteralPath $pdfPath) {
            Write-Host "PDF report written to:  $pdfPath" -ForegroundColor Green
        } else {
            Write-Warning "PDF was not created. You can generate it manually:"
            Write-Host "  `"$browserExe`" --headless=new --disable-gpu --no-pdf-header-footer --print-to-pdf=`"$pdfPath`" `"$fileUri`""
        }
    } catch {
        Write-Warning "PDF conversion failed: $($_.Exception.Message)"
        Write-Host "  `"$browserExe`" --headless=new --disable-gpu --no-pdf-header-footer --print-to-pdf=`"$pdfPath`" `"$fileUri`""
    }
} else {
    $fileUri = ([System.Uri](Resolve-Path -LiteralPath $htmlPath).Path).AbsoluteUri
    Write-Warning 'Microsoft Edge or Google Chrome was not found. PDF was not generated.'
    Write-Host 'Install Edge or Chrome and run:' -ForegroundColor Yellow
    Write-Host "  msedge.exe --headless=new --disable-gpu --no-pdf-header-footer --print-to-pdf=`"$pdfPath`" `"$fileUri`""
}

Write-Host ''
Write-Host '==================== Report Generation Complete ====================' -ForegroundColor Green
Write-Host ("Customer : {0}" -f $CustomerName)
Write-Host ("HTML     : {0}" -f $htmlPath)
Write-Host ("PDF      : {0}" -f $pdfPath)
Write-Host '===================================================================='
