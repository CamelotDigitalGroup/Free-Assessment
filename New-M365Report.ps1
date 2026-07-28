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
# Palette  (Camelot Digital Group brand: deep green + heraldic gold)
# ---------------------------------------------------------------------------
$COLOR_BG       = '#0d0d0d'
$COLOR_TEXT     = '#ffffff'
$COLOR_GRAY     = '#9ca3af'
$COLOR_GREEN    = '#22c55e'   # semantic PASS / positive
$COLOR_RED      = '#ef4444'   # semantic FAIL / negative
$COLOR_AMBER    = '#f59e0b'   # semantic WARN
$COLOR_BLUE     = '#D4A843'   # brand accent (gold) – replaces former blue accents
$COLOR_BRAND    = '#1B4332'   # Camelot deep green (decorative arcs / corners)
$COLOR_GOLD     = '#D4A843'   # Camelot heraldic gold

# ---------------------------------------------------------------------------
# Load JSON
# ---------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $JsonPath)) {
    throw "JSON file not found: $JsonPath"
}
Write-Host "Reading assessment data from: $JsonPath" -ForegroundColor Cyan
$data = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json

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
        'low'      { return "<span class='badge' style='border:1px solid #6b7280;color:#6b7280;'>low</span>" }
        default    { return "<span class='badge' style='border:1px solid #6b7280;color:#6b7280;'>$(Enc $sev)</span>" }
    }
}

# ---------------------------------------------------------------------------
# User / Device health metric icons (70px, stroke based)
# ---------------------------------------------------------------------------
function Get-MetricIcon {
    param([string]$Kind)
    $st = "fill='none' stroke='$COLOR_GRAY' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'"
    switch ($Kind) {
        'person'   { return "<svg width='40' height='40' viewBox='0 0 24 24'><circle cx='12' cy='8' r='4' $st/><path d='M4 21c0-4.4 3.6-7 8-7s8 2.6 8 7' $st/></svg>" }
        'admin'    { return "<svg width='40' height='40' viewBox='0 0 24 24'><path d='M6 10 L6 6 L9 8 L12 4 L15 8 L18 6 L18 10 Z' $st/><path d='M6 12 h12' $st/><path d='M5 20c0-3.3 3.1-5 7-5s7 1.7 7 5' $st/></svg>" }
        'lock'     { return "<svg width='40' height='40' viewBox='0 0 24 24'><circle cx='9' cy='8' r='3.5' $st/><path d='M2.5 20c0-3.3 2.9-5 6.5-5 1 0 2 .1 2.8.4' $st/><rect x='13' y='13' width='9' height='7' rx='1.2' $st/><path d='M15 13 v-2 a2.5 2.5 0 0 1 5 0 v2' $st/></svg>" }
        'lock-open'{ return "<svg width='40' height='40' viewBox='0 0 24 24'><rect x='5' y='11' width='14' height='10' rx='1.5' $st/><path d='M8 11 V7 a4 4 0 0 1 7.5-2' $st/><circle cx='12' cy='16' r='1.3' fill='$COLOR_GRAY' stroke='none'/></svg>" }
        'risky'    { return "<svg width='40' height='40' viewBox='0 0 24 24'><circle cx='10' cy='8' r='3.5' $st/><path d='M3 20c0-3.6 3.1-6 7-6' $st/><line x1='15' y1='14' x2='21' y2='20' $st/><line x1='21' y1='14' x2='15' y2='20' $st/></svg>" }
        'devices'  { return "<svg width='40' height='40' viewBox='0 0 24 24'><rect x='2' y='4' width='14' height='10' rx='1' $st/><path d='M6 18 h6' $st/><path d='M9 14 v4' $st/><rect x='16' y='9' width='6' height='11' rx='1.2' $st/></svg>" }
        'doc-alert'{ return "<svg width='40' height='40' viewBox='0 0 24 24'><path d='M7 3 h7 l5 5 v13 a1 1 0 0 1-1 1 H7 a1 1 0 0 1-1-1 V4 a1 1 0 0 1 1-1 Z' $st/><path d='M14 3 v5 h5' $st/><line x1='12.5' y1='11' x2='12.5' y2='15.5' $st/><circle cx='12.5' cy='18' r='.9' fill='$COLOR_GRAY' stroke='none'/></svg>" }
        'calendar' { return "<svg width='40' height='40' viewBox='0 0 24 24'><rect x='3' y='5' width='18' height='16' rx='1.5' $st/><path d='M3 9 h18' $st/><path d='M8 3 v4 M16 3 v4' $st/><line x1='12' y1='12' x2='12' y2='16' $st/><circle cx='12' cy='18.3' r='.9' fill='$COLOR_GRAY' stroke='none'/></svg>" }
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
# Decorative SVG (cover arc + summary corner)
# ---------------------------------------------------------------------------
$CoverArc = @"
<svg class='cover-arc' width='700' height='1100' viewBox='0 0 700 1100' preserveAspectRatio='none'>
  <circle cx='-40' cy='550' r='560' fill='none' stroke='$COLOR_BRAND' stroke-width='60' opacity='0.95'/>
  <circle cx='-40' cy='550' r='560' fill='none' stroke='$COLOR_GOLD' stroke-width='3' opacity='0.5'/>
</svg>
"@

$SummaryArc = @"
<svg class='summary-arc' width='520' height='900' viewBox='0 0 520 900' preserveAspectRatio='none'>
  <circle cx='-260' cy='450' r='430' fill='none' stroke='$COLOR_BRAND' stroke-width='55' opacity='0.9'/>
</svg>
<div class='corner-rect'></div>
"@

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
    <circle cx='$cx' cy='$cy' r='$r' fill='none' stroke='#2a2a2a' stroke-width='14'
            stroke-dasharray='$trackLen $gap' stroke-linecap='round'/>
    <circle cx='$cx' cy='$cy' r='$r' fill='none' stroke='$COLOR_GREEN' stroke-width='14'
            stroke-dasharray='$fillLen $rest' stroke-linecap='round'/>
  </g>
  <text x='$cx' y='$($cy+2)' text-anchor='middle' fill='#ffffff' font-size='38' font-weight='700'>$p%</text>
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
        [void]$sb.Append("<line x1='$padL' y1='$gy' x2='$($W-$padR)' y2='$gy' stroke='#2a2a2a' stroke-width='1'/>")
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
    # Built-in Camelot Digital Group shield + wordmark
    $shield = @"
<svg class='cdg-shield' width='60' height='60' viewBox='0 0 256 256'>
  <path d='M128 16L36 54v82c0 58 92 100 92 100s92-42 92-100V54L128 16z' fill='$COLOR_BRAND' stroke='$COLOR_GOLD' stroke-width='6'/>
  <path d='M128 40L60 68v66c0 45 68 82 68 82s68-37 68-82V68L128 40z' fill='none' stroke='$COLOR_GOLD' stroke-width='3' opacity='0.7'/>
  <path d='M92 108 l-8-16 h12 l6 10 6-14 6 14 6-10 h12 l-8 16 z' fill='$COLOR_GOLD'/>
  <rect x='86' y='112' width='84' height='10' rx='2' fill='$COLOR_GOLD'/>
  <text x='128' y='168' font-family='Georgia, serif' font-size='58' font-weight='700' text-anchor='middle' fill='$COLOR_GOLD'>CDG</text>
</svg>
"@
    return "<div class='logo'>$shield<div class='logo-word'><span class='logo-text'>CAMELOT DIGITAL <span style='color:$COLOR_GOLD;'>GROUP</span></span><div class='logo-sub'>WHERE HERITAGE MEETS INNOVATION</div></div></div>"
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
        [void]$sb.Append("<polygon points='$($pts -join ' ')' fill='none' stroke='#2a2a2a' stroke-width='1'/>")
    }
    # spokes + labels
    for ($i = 0; $i -lt $n; $i++) {
        $ang = (-90 + 360.0 * $i / $n) * [math]::PI / 180.0
        $x = [math]::Round($cx + $R * [math]::Cos($ang), 1)
        $y = [math]::Round($cy + $R * [math]::Sin($ang), 1)
        [void]$sb.Append("<line x1='$cx' y1='$cy' x2='$x' y2='$y' stroke='#2a2a2a' stroke-width='1'/>")
        $lx = [math]::Round($cx + ($R + 26) * [math]::Cos($ang), 1)
        $ly = [math]::Round($cy + ($R + 26) * [math]::Sin($ang), 1)
        $anchor = 'middle'
        if ($lx -gt $cx + 10) { $anchor = 'start' } elseif ($lx -lt $cx - 10) { $anchor = 'end' }
        $lbl = Enc ($Axes[$i].Label)
        [void]$sb.Append("<text x='$lx' y='$($ly+4)' text-anchor='$anchor' fill='#e5e7eb' font-size='12'>$lbl</text>")
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
  <path d='M $($cx-$r) $cy A $r $r 0 0 1 $($cx+$r) $cy' fill='none' stroke='#2a2a2a' stroke-width='14' stroke-linecap='round'
        stroke-dasharray='$circR $circR'/>
  <path d='M $($cx-$r) $cy A $r $r 0 0 1 $($cx+$r) $cy' fill='none' stroke='$COLOR_GOLD' stroke-width='14' stroke-linecap='round'
        stroke-dasharray='$fill $rest'/>
  <text x='$cx' y='$($cy-6)' text-anchor='middle' fill='#ffffff' font-size='30' font-weight='700'>$Value/$Total</text>
  <text x='$cx' y='$($cy+16)' text-anchor='middle' fill='$COLOR_GOLD' font-size='13' font-weight='600' letter-spacing='1'>$(Enc $Label)</text>
</svg>
"@
}

# ===========================================================================
# BUILD PAGES
# ===========================================================================
$pages = New-Object System.Text.StringBuilder

# ---- Page 1: Cover -------------------------------------------------------
$logoHtml = Get-LogoHtml
[void]$pages.Append(@"
<section class='page cover'>
  $CoverArc
  <div class='cover-top'>$logoHtml</div>
  <div class='cover-body'>
    <div class='cover-date'>$(Enc $reportDateStr)</div>
    <h1 class='cover-name'>$(Enc $CustomerName)</h1>
    <h2 class='cover-title'>Cloud Assessment Report</h2>
    <div class='cover-rule'></div>
    <div class='cover-exec'>Executive Summary</div>
  </div>
  <div class='cover-foot'>Prepared by $(Enc $PreparedBy)</div>
</section>
"@)

# ---- Page 2: Secure Score ------------------------------------------------
$ss = $data.SecureScore
$ssPct = 0; $ssCur = 0.0; $ssMax = 0.0
if ($ss) {
    try { $ssCur = [double]$ss.Current } catch {}
    try { $ssMax = [double]$ss.Max } catch {}
    if ($ss.Percentage -ne $null) { try { $ssPct = [int]$ss.Percentage } catch {} }
    elseif ($ssMax -gt 0) { $ssPct = [int][math]::Round(($ssCur/$ssMax)*100) }
}
$gauge = Get-GaugeSvg -Percentage $ssPct -Current $ssCur -Max $ssMax
$chart = Get-LineChartSvg -History ($ss.History)
$tcs = [math]::Max(0, [math]::Min(100, $TypicalClientScore))
$yourPos = [math]::Max(0, [math]::Min(100, $ssPct))

[void]$pages.Append(@"
<section class='page'>
  $SummaryArc
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

# ---- Page 3: User Health -------------------------------------------------
$uh = $data.UserHealth
$totalUsers = [int]($uh.TotalUsers)
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
  $SummaryArc
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
  $SummaryArc
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
  $SummaryArc
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
# SharePoint Site Inventory (paginated) — lists ALL discovered sites
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
  $SummaryArc
  <div class='corner-rect'></div>
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
  $SummaryArc
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
            $barColor = if ($pct -gt 90) { '$COLOR_RED' } elseif ($pct -gt 70) { '#f59e0b' } else { '$COLOR_GREEN' }
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
    $licRowsHtml = "<tr><td colspan='4' style='color:#9ca3af;padding:12px;text-align:center;'>No licensing data was available. Ensure <code>Organization.Read.All</code> is granted.</td></tr>"
}

[void]$pages.Append(@"
<section class='page'>
  $SummaryArc
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
        }
        $amCardsHtml = ($methods | ForEach-Object {
            $mid   = [string]$_.Method
            $state = [string]$_.State
            $fname = if ($friendlyNames.ContainsKey($mid)) { $friendlyNames[$mid] } else { $mid -replace '([a-z])([A-Z])','$1 $2' }
            $cls   = switch ($state.ToLower()) { 'enabled' { 'am-enabled' }; 'disabled' { 'am-disabled' }; default { 'am-other' } }
            "<div class='am-card'><div class='am-name'>$(Enc $fname)</div><div class='am-state $cls'>$(Enc $state)</div></div>"
        }) -join "`n"
    }
}
if ([string]::IsNullOrWhiteSpace($amCardsHtml)) {
    $amCardsHtml = "<p style='color:#9ca3af;text-align:center;padding:16px;'>Authentication methods data was not available. Ensure <code>Policy.Read.All</code> is granted.</p>"
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
  $SummaryArc
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
  $SummaryArc
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
  $SummaryArc
  <h1 class='page-h1'>Security Baseline Overview</h1>
  <p class='body-text'>Security controls are grouped into four maturity tiers. Each gauge shows how many controls in that tier are currently passing out of the controls that could be evaluated. "Not Set" indicates controls that require manual review or were not configured.</p>
  <div class='tier-grid'>
    $($tierCardsHtml.ToString())
  </div>
  <p class='mf-note'>Controls requiring manual verification are reported as "Not Set" until reviewed.</p>
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

    $groupsHtml = New-Object System.Text.StringBuilder
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

        [void]$groupsHtml.Append("<div class='check-group'><div class='check-head'><div class='check-icon'>$pIcon</div><div class='check-title'>$(Enc $g.Name) - $(Enc $parentTitle)</div></div><div class='check-subs'>$($subs.ToString())</div></div>")
    }

    [void]$pages.Append(@"
<section class='page section-page'>
  <h1 class='section-h1'>$(Enc $secTitle)</h1>
  <div class='check-list'>
    $($groupsHtml.ToString())
  </div>
</section>
"@)
}

# ===========================================================================
# CSS
# ===========================================================================
$css = @'
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&family=Inter:wght@300;400;600;700&display=swap');
* { margin:0; padding:0; box-sizing:border-box; print-color-adjust:exact; -webkit-print-color-adjust:exact; }
html, body { background:#000; font-family:'Inter', 'Segoe UI', Arial, Helvetica, sans-serif; color:#ffffff; }
.page {
  width:210mm; min-height:297mm; background:#0d0d0d;
  padding:48px 56px; page-break-after:always; position:relative; overflow:hidden;
}
.page:last-child { page-break-after:auto; }

/* Decorations */
.cover-arc { position:absolute !important; left:-120px; top:0; height:100%; z-index:0; }
.summary-arc { position:absolute !important; left:-160px; bottom:-80px; z-index:0; }
.corner-rect { position:absolute !important; right:-30px; top:-30px; width:150px; height:70px; background:#1B4332; opacity:0.6; border-radius:0 0 0 60px; z-index:0; }
.page > * { position:relative; z-index:1; }

/* Cover */
.cover { display:flex; flex-direction:column; }
.logo { display:flex; align-items:center; gap:14px; }
.cdg-shield { flex:0 0 auto; }
.logo-word { display:flex; flex-direction:column; }
.logo-text { font-family:'Playfair Display', Georgia, serif; font-size:24px; font-weight:700; letter-spacing:2px; color:#ffffff; }
.logo-sub { font-size:9px; letter-spacing:3px; color:#D4A843; margin-top:3px; }
.logo img, .logo svg { max-height:60px; }
.cover-top { margin-bottom:120px; }
.cover-body { margin-top:60px; }
.cover-date { font-size:20px; color:#D4A843; margin-bottom:40px; letter-spacing:1px; }
.cover-name { font-family:'Playfair Display', Georgia, serif; font-size:70px; font-weight:700; line-height:1.02; color:#ffffff; }
.cover-title { font-family:'Playfair Display', Georgia, serif; font-size:56px; font-weight:400; line-height:1.05; color:#ffffff; margin-top:2px; }
.cover-rule { height:2px; background:linear-gradient(to right,#D4A843,rgba(212,168,67,0)); width:92%; margin:34px 0 26px; }
.cover-exec { font-size:24px; color:#d1d5db; font-weight:300; }
.cover-foot { margin-top:auto; color:#9ca3af; font-size:12px; letter-spacing:1px; }

/* Headings */
.page-h1 { font-family:'Playfair Display', Georgia, serif; font-size:50px; font-weight:400; color:#ffffff; margin-bottom:34px; }
.section-h1 { font-family:'Playfair Display', Georgia, serif; font-size:26px; font-weight:700; color:#ffffff; margin-bottom:28px; }
.body-text { font-size:15px; line-height:1.55; color:#e5e7eb; margin:20px 0; max-width:95%; }

/* Secure score */
.ss-row { display:flex; align-items:center; gap:30px; margin:30px 0 10px; }
.ss-gauge { flex:0 0 auto; }
.ss-chart { flex:1 1 auto; }
.cmp-wrap { display:flex; align-items:center; gap:18px; margin-top:60px; }
.cmp-label-left, .cmp-label-right { font-weight:700; font-size:14px; color:#ffffff; text-align:center; white-space:nowrap; }
.cmp-track-wrap { position:relative; flex:1 1 auto; padding:34px 0; }
.cmp-track { height:16px; border-radius:8px; background:linear-gradient(to right,#ec1e63,#ef4444,#f59e0b,#eab308,#84cc16,#22c55e); position:relative; }
.cmp-dot { position:absolute; top:50%; width:20px; height:20px; border-radius:50%; background:#d1d5db; border:2px solid #0d0d0d; transform:translate(-50%,-50%); box-shadow:0 0 0 1px #9ca3af; }
.cmp-marker-top, .cmp-marker-bot { position:absolute; transform:translateX(-50%); }
.cmp-marker-top { top:0; } .cmp-marker-bot { bottom:0; }
.cmp-tip { display:inline-block; background:#e5e7eb; color:#111827; font-size:12px; font-weight:600; padding:4px 9px; border-radius:5px; white-space:nowrap; }

/* Metric rows (user/device health) */
.metric-list { display:flex; flex-direction:column; gap:34px; margin-top:10px; }
.metric-row { display:grid; grid-template-columns:110px 260px 1fr; align-items:center; gap:8px; }
.metric-circle { width:88px; height:88px; border-radius:50%; background:#1c1c1c; display:flex; align-items:center; justify-content:center; }
.metric-num { font-size:44px; font-weight:600; color:#ffffff; display:flex; align-items:center; gap:8px; line-height:1; }
.metric-ind { display:inline-flex; }
.metric-label { font-size:15px; color:#e5e7eb; margin-top:8px; line-height:1.3; }
.metric-desc { font-size:14px; color:#e5e7eb; line-height:1.5; }

/* Applications & Data */
.ad-sub { font-size:20px; font-weight:700; color:#ffffff; margin:28px 0 14px; }
.share-row { display:flex; align-items:flex-start; gap:16px; }
.share-icon { flex:0 0 auto; margin-top:2px; }
.share-title { font-size:16px; font-weight:600; color:#ffffff; }
.share-desc { font-size:14px; color:#9ca3af; margin-top:3px; max-width:90%; }
.sp-list { display:flex; flex-direction:column; gap:16px; margin-top:6px; }
.sp-row { display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid #1f1f1f; padding-bottom:10px; }
.sp-name { font-size:15px; font-weight:700; color:#ffffff; }
.sp-count-num { font-size:18px; color:#ffffff; }
.sp-count-lbl { font-size:12px; color:#9ca3af; }
.sp-empty { color:#9ca3af; font-size:14px; }

/* SharePoint site inventory table */
.spi-page { font-size:14px; color:#9ca3af; font-weight:400; }
.spi-table { width:100%; border-collapse:collapse; margin-top:14px; font-size:12.5px; }
.spi-table thead th { text-align:left; color:#D4A843; font-weight:700; border-bottom:1px solid #333; padding:8px 10px; font-size:12px; text-transform:uppercase; letter-spacing:.4px; }
.spi-table tbody td { padding:7px 10px; border-bottom:1px solid #1a1a1a; color:#e5e7eb; vertical-align:top; }
.spi-table tbody tr:nth-child(even) td { background:#141414; }
.spi-name { font-weight:600; color:#ffffff; width:38%; word-break:break-word; }
.spi-url { color:#9ca3af; word-break:break-all; }

/* Section check pages */
.section-page { background:#000000; }
.check-list { display:flex; flex-direction:column; gap:22px; }
.check-group { }
.check-head { display:flex; align-items:flex-start; gap:14px; }
.check-icon { flex:0 0 auto; margin-top:1px; }
.check-title { font-size:15px; font-weight:700; color:#ffffff; line-height:1.35; padding-top:6px; }
.check-subs { margin:10px 0 0 54px; display:flex; flex-direction:column; gap:14px; }
.subcheck { display:flex; align-items:flex-start; gap:12px; }
.sub-icon { flex:0 0 auto; margin-top:1px; }
.sub-title { font-size:14px; font-weight:600; color:#ffffff; line-height:1.3; }
.sub-detail { margin-top:4px; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
.sub-detail-text { font-size:12px; color:#9ca3af; line-height:1.4; }
.badge { font-size:10px; padding:1px 6px; border-radius:3px; font-weight:600; display:inline-block; line-height:1.5; }

/* Email Health */
.eh-table { width:100%; border-collapse:collapse; margin:14px 0 10px; }
.eh-table thead th { font-size:12px; font-weight:700; color:#D4A843; text-transform:uppercase; letter-spacing:1px; text-align:center; padding:8px 6px; border-bottom:2px solid #2a2a2a; }
.eh-table thead th.eh-dom { text-align:left; }
.eh-table td { padding:7px 6px; border-bottom:1px solid #1a1a1a; font-size:13px; vertical-align:middle; }
.eh-table td.eh-dom { text-align:left; color:#ffffff; font-weight:600; }
.eh-table td.eh-c { text-align:center; }
.eh-table td.eh-c svg { vertical-align:middle; }
.eh-default { color:#22c55e; font-weight:700; }
.eh-no { color:#9ca3af; }
.badge-pass { display:inline-block; padding:2px 8px; border-radius:4px; font-size:10px; font-weight:700; letter-spacing:1px; background:#1B4332; color:#D4A843; }
.badge-fail { display:inline-block; padding:2px 8px; border-radius:4px; font-size:10px; font-weight:700; letter-spacing:1px; background:#dc2626; color:#ffffff; }
.badge-na   { display:inline-block; padding:2px 8px; border-radius:4px; font-size:10px; font-weight:700; letter-spacing:1px; background:#374151; color:#9ca3af; }

/* Licensing table */
.lic-table { width:100%; border-collapse:collapse; margin:14px 0 10px; }
.lic-table thead th { font-size:11px; font-weight:700; color:#D4A843; text-transform:uppercase; letter-spacing:1px; text-align:left; padding:8px 6px; border-bottom:2px solid #2a2a2a; }
.lic-table td { padding:7px 6px; border-bottom:1px solid #1a1a1a; font-size:12px; vertical-align:middle; color:#ffffff; }
.lic-table td.lic-name { font-weight:600; max-width:240px; overflow:hidden; text-overflow:ellipsis; }
.lic-bar-track { width:120px; height:8px; background:#1a1a1a; border-radius:4px; display:inline-block; vertical-align:middle; margin-right:8px; }
.lic-bar-fill { height:100%; border-radius:4px; }
.lic-pct { font-size:11px; color:#9ca3af; }

/* Auth methods */
.am-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:8px; margin:14px 0; }
.am-card { background:#111; border:1px solid #222; border-radius:8px; padding:12px 10px; text-align:center; }
.am-name { font-size:11px; color:#d1d5db; font-weight:600; text-transform:capitalize; }
.am-state { font-size:10px; font-weight:700; margin-top:4px; text-transform:uppercase; letter-spacing:1px; }
.am-enabled  { color:#22c55e; }
.am-disabled { color:#ef4444; }
.am-other    { color:#9ca3af; }

/* Guest access */
.ga-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:12px; margin:14px 0; }
.ga-card { background:#111; border:1px solid #222; border-radius:8px; padding:16px; text-align:center; }
.ga-num { font-size:32px; font-weight:800; color:#D4A843; line-height:1; }
.ga-label { font-size:11px; color:#9ca3af; margin-top:6px; }
.eh-pol { display:inline-block; margin-left:6px; font-size:11px; color:#9ca3af; vertical-align:middle; }
.eh-empty { text-align:center; color:#9ca3af; padding:20px; }
.mf-bars { display:flex; flex-direction:column; gap:16px; margin-top:8px; }
.mf-item { }
.mf-top { display:flex; justify-content:space-between; align-items:baseline; margin-bottom:6px; }
.mf-label { font-size:14px; color:#e5e7eb; }
.mf-num { font-size:16px; font-weight:700; color:#ffffff; }
.mf-track { height:14px; background:#1c1c1c; border-radius:7px; overflow:hidden; }
.mf-fill { height:100%; border-radius:7px; }
.mf-note { font-size:12px; color:#9ca3af; margin-top:14px; }

/* Radar */
.radar-wrap { display:flex; flex-direction:column; align-items:center; margin-top:6px; }
.radar-legend { margin-top:10px; font-size:13px; color:#e5e7eb; display:flex; align-items:center; gap:8px; }
.radar-swatch { display:inline-block; width:22px; height:12px; background:#D4A843; opacity:0.6; border:1px solid #D4A843; border-radius:2px; }

/* Baseline overview tiers */
.tier-grid { display:grid; grid-template-columns:1fr 1fr; gap:36px 28px; margin-top:20px; }
.tier-card { display:flex; align-items:center; gap:20px; background:#141414; border:1px solid #222; border-radius:12px; padding:18px 22px; }
.tier-legend { display:flex; flex-direction:column; gap:7px; }
.tl-row { display:flex; align-items:center; gap:9px; font-size:14px; color:#e5e7eb; }
.tl-ico { display:inline-block; width:14px; height:14px; border-radius:50%; flex:0 0 auto; }
.tl-ico.ok { background:#22c55e; }
.tl-ico.bad { background:#ef4444; }
.tl-ico.amber { background:#f59e0b; }
.tl-ico.gray { background:#6b7280; }
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
@page { size:A4 portrait; margin:0; }
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
