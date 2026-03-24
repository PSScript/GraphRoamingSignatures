<#
.SYNOPSIS
    Set-GraphSignature.ps1 — Email Signature Deployment via Microsoft Graph
    Roaming Signatures • Delta-Aware • Multi-Company • VIP Overrides

.DESCRIPTION
    Deploys HTML email signatures to Exchange Online mailboxes using the
    Graph Beta UserConfiguration API (MailboxConfigItem.ReadWrite).

    Change detection via SHA256 checksum stored in a configurable extensionAttribute.
    Graph /users/delta narrows processing to only changed users.

    Assignment hierarchy (first match wins):
      1. VIP override       (UPN → personal template)
      2. Security group     (group membership → group template)
      3. companyName match  (Entra companyName → company template)
      4. Default fallback

    Template variables use %Variable% syntax — same as Exchange transport rules.

.PARAMETER Force
    Ignore checksums, redeploy ALL users. Use after template changes.

.PARAMETER WhatIf
    Preview mode — show what would change without deploying.

.PARAMETER UserUPN
    Deploy to a single user only. Omit for bulk run.

.EXAMPLE
    # Single user test
    .\Set-GraphSignature.ps1 -UserUPN "j.doe@contoso.com" -WhatIf

    # Scheduled bulk run
    .\Set-GraphSignature.ps1

    # Force full redeploy after template change
    .\Set-GraphSignature.ps1 -Force
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$UserUPN,
    [switch]$Force,
    [string]$StatePath = '.\signature-state.json',
    [string]$LogPath   = ".\logs\sig-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web

# ═══════════════════════════════════════════════════════════════════════════
# CONFIGURATION — edit this section for your environment
# ═══════════════════════════════════════════════════════════════════════════

# ── Entra ID / App Registration ──────────────────────────────────────────
$TenantId     = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
$ClientId     = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
$ClientSecret = 'your-client-secret-here'

# ── Domain & Branding ────────────────────────────────────────────────────
$PrimaryDomain = 'contoso.com'
$CompanyUrl    = "https://www.$PrimaryDomain"

# Logo / photo base URL (must be publicly reachable or embedded as base64)
$AssetBaseUrl  = "https://assets.$PrimaryDomain/signatures"

# ── Checksum tracking ────────────────────────────────────────────────────
# Which extensionAttribute to store the change-detection hash in.
# Must be cloud-writeable (not synced from on-prem AD via AD Connect).
$ChecksumAttribute = 'extensionAttribute15'
$ChecksumPrefix    = 'SIG:'

# ── Throttling ───────────────────────────────────────────────────────────
$ThrottleDelayMs = 100   # Delay between Graph calls (ms)
$BatchSize       = 20    # Unused currently, reserved for parallel runs

# ── Signature fields ─────────────────────────────────────────────────────
# These are the user properties pulled from Graph and available as
# %Variable% tokens in templates. Add or remove as needed.
$SignatureFields = @(
    'DisplayName'
    'GivenName'
    'Surname'
    'JobTitle'
    'Department'
    'Mail'
    'Phone'              # → businessPhones[0]
    'Mobile'             # → mobilePhone
    'Fax'                # → faxNumber
    'Company'            # → companyName
    'Office'             # → officeLocation
    'Street'             # → streetAddress
    'City'
    'PostalCode'
    'State'
    'Country'
    'MailNickname'
    'ManagerName'        # → manager.displayName
    'ManagerMail'        # → manager.mail
    'ExtensionAttribute1'
    'ExtensionAttribute2'
    'ExtensionAttribute3'
    'ExtensionAttribute4'
    'ExtensionAttribute5'
    'ExtensionAttribute6'
    'ExtensionAttribute7'
    'ExtensionAttribute8'
    'ExtensionAttribute9'
    'ExtensionAttribute10'
    'ExtensionAttribute11'
    'ExtensionAttribute12'
    'ExtensionAttribute13'
    'ExtensionAttribute14'
    'ExtensionAttribute15'
)

# ── Departments (informational — used for reporting, not assignment) ─────
$Departments = @(
    'Engineering'
    'Development'
    'Sales'
    'Marketing'
    'People 'Human Resources' Culture'
    'Finance'
    'Legal'
    'Operations'
    'Management'
)

# ── Company → Template mapping ───────────────────────────────────────────
# Key = Entra ID companyName (exact match, case-insensitive)
# Value = hashtable with template HTML and signature display name
$CompanyTemplates = @{
    'Contoso Group' = @{
        SignatureName = 'Contoso Group Signature'
        LogoUrl       = "$AssetBaseUrl/logo-contoso.png"
        AccentColor   = '#E30613'
    }
    'Contoso Cloud' = @{
        SignatureName = 'Contoso Cloud Signature'
        LogoUrl       = "$AssetBaseUrl/logo-contoso-cloud.png"
        AccentColor   = '#E30613'
    }
    'Contoso Solutions' = @{
        SignatureName = 'Contoso Solutions Signature'
        LogoUrl       = "$AssetBaseUrl/logo-contoso-solutions.png"
        AccentColor   = '#E30613'
    }
}

# ── Default template (fallback when companyName doesn't match) ───────────
$DefaultTemplate = @{
    SignatureName = 'Contoso Group Signature'
    LogoUrl       = "$AssetBaseUrl/logo-contoso.png"
    AccentColor   = '#E30613'
}

# ── VIP overrides (UPN → custom template config) ────────────────────────
# These users get a personal template regardless of company/group.
# Use for CEOs, managing directors, etc.
$VipOverrides = @{
    "ceo@$PrimaryDomain" = @{
        SignatureName = 'CEO Signature'
        LogoUrl       = "$AssetBaseUrl/logo-contoso.png"
        PhotoUrl      = "$AssetBaseUrl/photos/ceo.jpg"
        AccentColor   = '#E30613'
        CustomHtml    = $null  # set to a full HTML string to override the standard layout entirely
        LinkedIn      = 'https://linkedin.com/in/ceo-contoso'
    }
    "md.north@$PrimaryDomain" = @{
        SignatureName = 'MD North Signature'
        LogoUrl       = "$AssetBaseUrl/logo-contoso.png"
        PhotoUrl      = "$AssetBaseUrl/photos/md-north.jpg"
        AccentColor   = '#E30613'
        LinkedIn      = 'https://linkedin.com/in/md-north'
    }
    "md.south@$PrimaryDomain" = @{
        SignatureName = 'MD South Signature'
        LogoUrl       = "$AssetBaseUrl/logo-contoso.png"
        PhotoUrl      = "$AssetBaseUrl/photos/md-south.jpg"
        AccentColor   = '#E30613'
        LinkedIn      = 'https://linkedin.com/in/md-south'
    }
}

# ── Security group overrides (groupId → template config) ────────────────
# Members of these groups get their group template (overrides company match).
$GroupOverrides = @{
    # 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' = @{
    #     DisplayName   = 'SG-Signature-Marketing'
    #     SignatureName = 'Marketing Campaign Q2'
    #     BannerUrl     = "$AssetBaseUrl/banners/campaign-q2.png"
    #     BannerWidth   = 480
    #     BannerHeight  = 96
    #     BannerLink    = "https://www.$PrimaryDomain/campaign"
    # }
}

# ── Campaign banner (optional, appended below signature) ─────────────────
# Set to $null to disable. Applied to ALL users unless overridden.
$CampaignBanner = $null
# $CampaignBanner = @{
#     ImageUrl = "$AssetBaseUrl/banners/campaign-2026-q2.png"
#     Width    = 480
#     Height   = 96
#     AltText  = 'Visit our booth at IT-SA 2026'
#     LinkUrl  = "https://www.$PrimaryDomain/events/itsa2026"
# }

# ── Exclude patterns ────────────────────────────────────────────────────
$ExcludeUPNs = @(
    "admin@$PrimaryDomain"
    "noreply@$PrimaryDomain"
    "postmaster@$PrimaryDomain"
    "servicedesk@$PrimaryDomain"
)

$ExcludeUPNPrefixes = @(
    'svc-'
    'shared-'
    'room-'
    'equipment-'
    'test-'
)

$ExcludeCompanyNames = @(
    'Service Accounts'
    'Test Accounts'
)


# ═══════════════════════════════════════════════════════════════════════════
# HTML TEMPLATES — %Variable% tokens replaced at deploy time
# ═══════════════════════════════════════════════════════════════════════════

function Get-StandardSignatureHtml {
    <#
    .SYNOPSIS
        Standard layout: contact info left, photo/logo right (240x160).
        Uses %Variable% tokens that get replaced with Graph user data.
        Accent color and logo injected from the template config.
    #>
    param(
        [string]$LogoUrl,
        [string]$AccentColor = '#E30613',
        [string]$PhotoUrl    = $null,
        [string]$LinkedIn    = $null,
        [string]$BannerUrl   = $null,
        [int]$BannerWidth    = 480,
        [int]$BannerHeight   = 96,
        [string]$BannerLink  = $null
    )

    # Right column: photo if available, otherwise logo
    $rightImageUrl    = if ($PhotoUrl) { $PhotoUrl } else { $LogoUrl }
    $rightImageWidth  = if ($PhotoUrl) { 240 } else { 160 }
    $rightImageHeight = if ($PhotoUrl) { 160 } else { 60 }
    $rightImageAlt    = if ($PhotoUrl) { '%DisplayName%' } else { '%Company%' }

    $linkedInRow = ''
    if ($LinkedIn) {
        $linkedInRow = @"
        <tr>
          <td style="padding-top:6px;font-size:9pt;">
            <a href="$LinkedIn" style="color:${AccentColor};text-decoration:none;">LinkedIn</a>
          </td>
        </tr>
"@
    }

    $bannerBlock = ''
    if ($BannerUrl) {
        $img = "<img src=`"$BannerUrl`" width=`"$BannerWidth`" height=`"$BannerHeight`" alt=`"Campaign`" style=`"display:block;border:0;outline:none;`" />"
        if ($BannerLink) {
            $img = "<a href=`"$BannerLink`" style=`"text-decoration:none;`">$img</a>"
        }
        $bannerBlock = @"
<table cellpadding="0" cellspacing="0" border="0" style="margin-top:12px;">
  <tr><td>$img</td></tr>
</table>
"@
    }

    $html = @"
<table cellpadding="0" cellspacing="0" border="0" style="font-family:'Segoe UI',Calibri,Arial,Helvetica,sans-serif;font-size:10pt;color:#333333;max-width:600px;">
  <tr>
    <!-- LEFT: Contact info -->
    <td style="padding-right:16px;vertical-align:top;">
      <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td style="font-size:13pt;font-weight:600;color:#1a1a1a;padding-bottom:2px;">
            %DisplayName%
          </td>
        </tr>
        <tr>
          <td style="font-size:9pt;color:${AccentColor};text-transform:uppercase;letter-spacing:0.5px;padding-bottom:8px;">
            %JobTitle%
          </td>
        </tr>
        <tr>
          <td style="font-size:9pt;color:#555555;line-height:1.8;">
            <span style="color:#999;">T</span>&nbsp; %Phone%<br/>
            <span style="color:#999;">M</span>&nbsp; %Mobile%<br/>
            <span style="color:#999;">E</span>&nbsp; <a href="mailto:%Mail%" style="color:${AccentColor};text-decoration:none;">%Mail%</a>
          </td>
        </tr>
        <tr>
          <td style="padding-top:8px;font-size:8pt;color:#999999;border-top:1px solid #e0e0e0;">
            %Company%<br/>
            %Street% &middot; %PostalCode% %City%
          </td>
        </tr>
        <tr>
          <td style="padding-top:4px;font-size:8pt;color:#bbbbbb;">
            %ExtensionAttribute7%
          </td>
        </tr>
${linkedInRow}
      </table>
    </td>
    <!-- RIGHT: Photo or Logo (240x160 / 160x60) -->
    <td style="vertical-align:top;padding-left:16px;border-left:3px solid ${AccentColor};">
      <img src="${rightImageUrl}" alt="${rightImageAlt}"
           width="${rightImageWidth}" height="${rightImageHeight}"
           style="display:block;border:0;outline:none;border-radius:4px;" />
    </td>
  </tr>
</table>
${bannerBlock}
"@

    return $html
}

function Get-VipSignatureHtml {
    <#
    .SYNOPSIS
        VIP layout: same structure but with personal photo (240x160)
        and optional LinkedIn link. Override entirely with CustomHtml if set.
    #>
    param([hashtable]$VipConfig)

    if ($VipConfig.CustomHtml) {
        return $VipConfig.CustomHtml
    }

    return Get-StandardSignatureHtml `
        -LogoUrl      $VipConfig.LogoUrl `
        -AccentColor  $VipConfig.AccentColor `
        -PhotoUrl     $VipConfig.PhotoUrl `
        -LinkedIn     $VipConfig.LinkedIn
}


# ═══════════════════════════════════════════════════════════════════════════
# ENCODING ENGINE
# ═══════════════════════════════════════════════════════════════════════════

function Get-StringHash {
    param([string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hash = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Text))
    return [BitConverter]::ToString($hash).Replace('-','').ToLower()
}


# ═══════════════════════════════════════════════════════════════════════════
# AUTHENTICATION
# ═══════════════════════════════════════════════════════════════════════════

function Get-GraphToken {
    $body = @{
        grant_type    = 'client_credentials'
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
    }
    $r = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/${TenantId}/oauth2/v2.0/token" `
        -ContentType 'application/x-www-form-urlencoded' -Body $body
    $script:Token = $r.access_token
    $script:TokenExpires = (Get-Date).AddSeconds($r.expires_in - 300)
}

function Invoke-Graph {
    param(
        [string]$Method = 'GET',
        [string]$Uri,
        [object]$Body,
        [int]$Retries = 3
    )

    if (-not $script:Token -or (Get-Date) -ge $script:TokenExpires) { Get-GraphToken }

    $base = 'https://graph.microsoft.com/beta'
    $fullUri = if ($Uri.StartsWith('http')) { $Uri } else { "${base}${Uri}" }
    $headers = @{
        Authorization    = "Bearer $($script:Token)"
        'Content-Type'   = 'application/json'
        ConsistencyLevel = 'eventual'
    }

    for ($r = 0; $r -lt $Retries; $r++) {
        try {
            $p = @{ Method=$Method; Uri=$fullUri; Headers=$headers }
            if ($Body -and $Method -ne 'GET') {
                $p.Body = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 -Compress }
            }
            return Invoke-RestMethod @p -ErrorAction Stop
        }
        catch {
            $code = $_.Exception.Response.StatusCode.Value__
            if ($code -eq 429) {
                $wait = 30
                try { $wait = [int]($_.Exception.Response.Headers | Where-Object Key -eq 'Retry-After').Value[0] } catch {}
                Write-Log "  Throttled 429 — waiting ${wait}s" -Level WARN
                Start-Sleep -Seconds $wait
            }
            elseif ($code -in 500,502,503,504 -and $r -lt $Retries-1) {
                Start-Sleep -Seconds ([math]::Pow(2,$r)*3)
            }
            else { throw }
        }
    }
}


# ═══════════════════════════════════════════════════════════════════════════
# LOGGING
# ═══════════════════════════════════════════════════════════════════════════

$script:Stats = @{ Processed=0; Deployed=0; Skipped=0; Errors=0; Start=Get-Date }

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','OK')]$Level='INFO')
    $ts = Get-Date -Format 'HH:mm:ss'
    $pre = switch($Level) { 'ERROR' {'!!'} 'WARN' {' >'} 'OK' {' +'} default {'  '} }
    $color = switch($Level) { 'ERROR'{[ConsoleColor]::Red} 'WARN'{[ConsoleColor]::Yellow} 'OK'{[ConsoleColor]::Green} default{[ConsoleColor]::Gray} }
    Write-Host "[$ts]$pre $Message" -ForegroundColor $color
    $null = New-Item -Path (Split-Path $LogPath) -ItemType Directory -Force -ErrorAction SilentlyContinue
    "[$ts][$Level] $Message" | Out-File -FilePath $LogPath -Append -Encoding utf8
}


# ═══════════════════════════════════════════════════════════════════════════
# DELTA QUERY — only changed users since last run
# ═══════════════════════════════════════════════════════════════════════════

function Get-ChangedUsers {
    param([hashtable]$State)

    $select = @(
        'id','userPrincipalName','displayName','givenName','surname','mail',
        'jobTitle','department','companyName','officeLocation',
        'businessPhones','mobilePhone','faxNumber',
        'streetAddress','city','state','postalCode','country',
        'onPremisesExtensionAttributes','mailNickname'
    ) -join ','

    $users = [System.Collections.Generic.List[object]]::new()
    $response = $null

    if ($State.deltaLink -and -not $Force -and -not $UserUPN) {
        Write-Log "Using delta token from last run"
        try { $response = Invoke-Graph -Uri $State.deltaLink }
        catch { Write-Log "Delta expired — full sync" -Level WARN }
    }

    if (-not $response) {
        if ($UserUPN) {
            Write-Log "Single user mode: $UserUPN"
            $u = Invoke-Graph -Uri "/users/${UserUPN}?`$select=${select}"
            return @{ Users = @($u); IsFullSync = $false }
        }
        Write-Log "Full user sync"
        $response = Invoke-Graph -Uri "/users/delta?`$select=${select}&`$top=200"
    }

    while ($response) {
        if ($response.value) { $users.AddRange($response.value) }
        if ($response.'@odata.deltaLink') {
            $State.deltaLink = $response.'@odata.deltaLink'
            $response = $null
        }
        elseif ($response.'@odata.nextLink') {
            Start-Sleep -Milliseconds $ThrottleDelayMs
            $response = Invoke-Graph -Uri $response.'@odata.nextLink'
        }
        else { $response = $null }
    }

    Write-Log "Delta returned $($users.Count) users"
    return @{ Users = $users; IsFullSync = (-not $State.deltaLink) }
}


# ═══════════════════════════════════════════════════════════════════════════
# TEMPLATE ASSIGNMENT — resolve which HTML a user gets
# ═══════════════════════════════════════════════════════════════════════════

$script:GroupMemberCache = @{}

function Resolve-Assignment {
    param([psobject]$User)

    $upn = $User.userPrincipalName

    # ── 1. VIP override ──────────────────────────────────────────────
    foreach ($vipUpn in $VipOverrides.Keys) {
        if ($upn -ieq $vipUpn) {
            $vc = $VipOverrides[$vipUpn]
            return @{
                Source        = "VIP:$vipUpn"
                SignatureName = $vc.SignatureName
                Html          = Get-VipSignatureHtml -VipConfig $vc
            }
        }
    }

    # ── 2. Security group override ───────────────────────────────────
    foreach ($groupId in $GroupOverrides.Keys) {
        $gc = $GroupOverrides[$groupId]

        # Lazy-load group members
        if (-not $script:GroupMemberCache.ContainsKey($groupId)) {
            try {
                $members = @()
                $resp = Invoke-Graph -Uri "/groups/${groupId}/members?`$select=id&`$top=999"
                while ($resp) {
                    $members += $resp.value.id
                    $resp = if ($resp.'@odata.nextLink') { Invoke-Graph -Uri $resp.'@odata.nextLink' } else { $null }
                }
                $script:GroupMemberCache[$groupId] = $members
            }
            catch {
                Write-Log "  Cannot fetch group ${groupId}: $_" -Level WARN
                $script:GroupMemberCache[$groupId] = @()
            }
        }

        if ($User.id -in $script:GroupMemberCache[$groupId]) {
            return @{
                Source        = "Group:$($gc.DisplayName)"
                SignatureName = $gc.SignatureName
                Html          = Get-StandardSignatureHtml `
                                    -LogoUrl $DefaultTemplate.LogoUrl `
                                    -AccentColor $DefaultTemplate.AccentColor `
                                    -BannerUrl $gc.BannerUrl `
                                    -BannerWidth $gc.BannerWidth `
                                    -BannerHeight $gc.BannerHeight `
                                    -BannerLink $gc.BannerLink
            }
        }
    }

    # ── 3. companyName match ─────────────────────────────────────────
    if ($User.companyName) {
        foreach ($compName in $CompanyTemplates.Keys) {
            if ($User.companyName -ieq $compName) {
                $ct = $CompanyTemplates[$compName]
                return @{
                    Source        = "Company:$compName"
                    SignatureName = $ct.SignatureName
                    Html          = Get-StandardSignatureHtml -LogoUrl $ct.LogoUrl -AccentColor $ct.AccentColor
                }
            }
        }
    }

    # ── 4. Default ───────────────────────────────────────────────────
    return @{
        Source        = 'Default'
        SignatureName = $DefaultTemplate.SignatureName
        Html          = Get-StandardSignatureHtml -LogoUrl $DefaultTemplate.LogoUrl -AccentColor $DefaultTemplate.AccentColor
    }
}


# ═══════════════════════════════════════════════════════════════════════════
# VARIABLE EXPANSION — replace %Variable% tokens with Graph user data
# ═══════════════════════════════════════════════════════════════════════════

function Expand-Variables {
    param([string]$Html, [psobject]$User, [psobject]$Manager)

    $r = $Html

    $map = [ordered]@{
        '%DisplayName%'    = $User.displayName
        '%GivenName%'      = $User.givenName
        '%Surname%'        = $User.surname
        '%Mail%'           = $User.mail
        '%UPN%'            = $User.userPrincipalName
        '%JobTitle%'       = $User.jobTitle
        '%Department%'     = $User.department
        '%Company%'        = $User.companyName
        '%Office%'         = $User.officeLocation
        '%Phone%'          = if ($User.businessPhones) { $User.businessPhones[0] } else { '' }
        '%Mobile%'         = $User.mobilePhone
        '%Fax%'            = $User.faxNumber
        '%Street%'         = $User.streetAddress
        '%City%'           = $User.city
        '%State%'          = $User.state
        '%PostalCode%'     = $User.postalCode
        '%Country%'        = $User.country
        '%MailNickname%'   = $User.mailNickname
        '%ManagerName%'    = if ($Manager) { $Manager.displayName } else { '' }
        '%ManagerMail%'    = if ($Manager) { $Manager.mail } else { '' }
    }

    # extensionAttribute1–15
    for ($i = 1; $i -le 15; $i++) {
        $val = ''
        if ($User.onPremisesExtensionAttributes) {
            $val = $User.onPremisesExtensionAttributes."extensionAttribute$i"
        }
        $map["%ExtensionAttribute$i%"] = $val
    }

    foreach ($key in $map.Keys) {
        $val = if ($null -eq $map[$key]) { '' } else { [System.Web.HttpUtility]::HtmlEncode($map[$key]) }
        $r = $r.Replace($key, $val)
    }

    # Clean unreplaced tokens
    $r = $r -replace '%\w+%', ''

    # Remove table rows that are now empty (all cells blank after variable replacement)
    $r = $r -replace '<tr>\s*(<td[^>]*>\s*(&nbsp;|\s|<br\s*/?>)*\s*</td>\s*)+</tr>', ''

    # Collapse excess whitespace
    $r = $r -replace '(\r?\n){3,}', "`r`n`r`n"

    # Append campaign banner if configured
    if ($CampaignBanner) {
        $img = "<img src=`"$($CampaignBanner.ImageUrl)`" width=`"$($CampaignBanner.Width)`" height=`"$($CampaignBanner.Height)`" alt=`"$($CampaignBanner.AltText)`" style=`"display:block;border:0;`" />"
        if ($CampaignBanner.LinkUrl) { $img = "<a href=`"$($CampaignBanner.LinkUrl)`" style=`"text-decoration:none;`">$img</a>" }
        $r += "`r`n<table cellpadding=`"0`" cellspacing=`"0`" border=`"0`" style=`"margin-top:12px;`"><tr><td>$img</td></tr></table>"
    }

    return $r
}

function ConvertTo-PlainText {
    param([string]$Html)
    $t = $Html -replace '<br\s*/?>', "`r`n" -replace '</p>', "`r`n" -replace '</div>', "`r`n" -replace '</tr>', "`r`n"
    $t = $t -replace '<[^>]+>', ''
    $t = [System.Web.HttpUtility]::HtmlDecode($t)
    $t = $t -replace '(\r?\n){3,}', "`r`n`r`n"
    return $t.Trim()
}


# ═══════════════════════════════════════════════════════════════════════════
# CHECKSUM — change detection in extensionAttribute
# ═══════════════════════════════════════════════════════════════════════════

function Get-UserChecksum {
    param([psobject]$User, [string]$TemplateHtml)

    $sig = @(
        $User.displayName, $User.givenName, $User.surname, $User.mail,
        $User.jobTitle, $User.department, $User.companyName, $User.officeLocation,
        ($User.businessPhones -join '|'), $User.mobilePhone, $User.faxNumber,
        $User.streetAddress, $User.city, $User.state, $User.postalCode, $User.country,
        (Get-StringHash $TemplateHtml).Substring(0,8)
    ) -join '|'

    return (Get-StringHash $sig).Substring(0, 12)
}

function Get-StoredChecksum {
    param([psobject]$User)
    $val = $null
    if ($User.onPremisesExtensionAttributes) {
        $attrNum = $ChecksumAttribute -replace 'extensionAttribute',''
        $val = $User.onPremisesExtensionAttributes."extensionAttribute$attrNum"
    }
    if ($val -and $val.StartsWith($ChecksumPrefix)) {
        return ($val.Substring($ChecksumPrefix.Length).Split('|'))[0]
    }
    return $null
}

function Set-StoredChecksum {
    param([string]$UserId, [string]$Hash)
    $value = "${ChecksumPrefix}${Hash}|$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ')"
    $attrNum = $ChecksumAttribute -replace 'extensionAttribute',''
    $body = @{ onPremisesExtensionAttributes = @{ "extensionAttribute$attrNum" = $value } }
    try {
        Invoke-Graph -Method PATCH -Uri "/users/${UserId}" -Body $body | Out-Null
        return $true
    }
    catch {
        if ("$_" -match 'DirSync|cloud-mastered') {
            Write-Log "    Cannot write checksum — attribute synced from on-prem AD" -Level WARN
        }
        return $false
    }
}


# ═══════════════════════════════════════════════════════════════════════════
# DEPLOYMENT — Graph Beta UserConfiguration API
# ═══════════════════════════════════════════════════════════════════════════

function Deploy-Signature {
    param(
        [string]$UserId,
        [string]$UPN,
        [string]$SignatureHtml,
        [string]$SignatureText,
        [string]$SignatureName
    )

    # inbox is a well-known folder name — no Mail.ReadWrite needed

    $dictXml = @"
<?xml version="1.0" encoding="utf-8"?>
<UserConfiguration>
  <Dictionary>
    <DictionaryEntry>
      <DictionaryKey><Type>String</Type><Value>signaturehtml</Value></DictionaryKey>
      <DictionaryValue><Type>String</Type><Value>$([System.Security.SecurityElement]::Escape($SignatureHtml))</Value></DictionaryValue>
    </DictionaryEntry>
    <DictionaryEntry>
      <DictionaryKey><Type>String</Type><Value>signaturetext</Value></DictionaryKey>
      <DictionaryValue><Type>String</Type><Value>$([System.Security.SecurityElement]::Escape($SignatureText))</Value></DictionaryValue>
    </DictionaryEntry>
    <DictionaryEntry>
      <DictionaryKey><Type>String</Type><Value>autoaddsignature</Value></DictionaryKey>
      <DictionaryValue><Type>Boolean</Type><Value>true</Value></DictionaryValue>
    </DictionaryEntry>
    <DictionaryEntry>
      <DictionaryKey><Type>String</Type><Value>autoaddsignatureonreply</Value></DictionaryKey>
      <DictionaryValue><Type>Boolean</Type><Value>true</Value></DictionaryValue>
    </DictionaryEntry>
  </Dictionary>
</UserConfiguration>
"@

    $base64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($dictXml))
    $body = @{ '@odata.type' = '#microsoft.graph.userConfiguration'; structuredData = $base64 }

    Invoke-Graph -Method PATCH `
        -Uri "/users/${UserId}/mailFolders/inbox/userConfigurations/OWA.UserOptions" `
        -Body $body | Out-Null
}


# ═══════════════════════════════════════════════════════════════════════════
# STATE PERSISTENCE
# ═══════════════════════════════════════════════════════════════════════════

function Load-State {
    if (Test-Path -LiteralPath $StatePath) {
        return Get-Content -LiteralPath $StatePath -Raw | ConvertFrom-Json -AsHashtable
    }
    return @{ deltaLink = $null; lastRun = $null }
}

function Save-State {
    param([hashtable]$State)
    $State.lastRun = Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ'
    $State | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $StatePath -Encoding utf8
}


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

function Main {
    Write-Host ""
    Write-Host "  Set-GraphSignature — Roaming Signature Deployment" -ForegroundColor Cyan
    Write-Host "  Domain: $PrimaryDomain | Companies: $($CompanyTemplates.Count) | VIPs: $($VipOverrides.Count)" -ForegroundColor DarkGray
    if ($Force) { Write-Host "  FORCE MODE — all users will be redeployed" -ForegroundColor Yellow }
    Write-Host ""

    $state = Load-State
    Write-Log "State loaded (last run: $($state.lastRun ?? 'never'))"

    Get-GraphToken
    Write-Log "Authenticated to $TenantId" -Level OK

    $delta = Get-ChangedUsers -State $state
    $users = $delta.Users

    # Filter
    $eligible = $users | Where-Object {
        $_.mail -and
        $_.userPrincipalName -and
        $_.userPrincipalName -notin $ExcludeUPNs -and
        $_.companyName -notin $ExcludeCompanyNames -and
        -not ($ExcludeUPNPrefixes | Where-Object { $_.userPrincipalName.StartsWith($_) })
    }

    $total = @($eligible).Count
    Write-Log "Eligible: $total (from $($users.Count) delta results)"

    $managerCache = @{}
    $i = 0

    foreach ($user in $eligible) {
        $i++
        $upn = $user.userPrincipalName
        $script:Stats.Processed++
        $pct = if ($total -gt 0) { [math]::Round($i/$total*100) } else { 0 }

        Write-Log "[$i/$total] ${pct}% $upn ($($user.companyName ?? '-'))"

        try {
            # Resolve template
            $assignment = Resolve-Assignment -User $user
            Write-Log "  → $($assignment.Source) [$($assignment.SignatureName)]"

            # Checksum
            $expected = Get-UserChecksum -User $user -TemplateHtml $assignment.Html
            $stored   = Get-StoredChecksum -User $user

            if (-not $Force -and $stored -eq $expected) {
                Write-Log "  Skip — checksum match ($expected)"
                $script:Stats.Skipped++
                continue
            }

            Write-Log "  Checksum: $($stored ?? 'none') → $expected — deploying"

            # Fetch manager (cached)
            if (-not $managerCache.ContainsKey($user.id)) {
                try { $managerCache[$user.id] = Invoke-Graph -Uri "/users/$($user.id)/manager?`$select=displayName,mail" }
                catch { $managerCache[$user.id] = $null }
            }

            # Expand variables
            $sigHtml = Expand-Variables -Html $assignment.Html -User $user -Manager $managerCache[$user.id]
            $sigText = ConvertTo-PlainText -Html $sigHtml

            # Deploy
            if ($PSCmdlet.ShouldProcess($upn, "Deploy '$($assignment.SignatureName)'")) {
                Deploy-Signature -UserId $user.id -UPN $upn `
                    -SignatureHtml $sigHtml -SignatureText $sigText `
                    -SignatureName $assignment.SignatureName

                $null = Set-StoredChecksum -UserId $user.id -Hash $expected
                $script:Stats.Deployed++
                Write-Log "  Deployed" -Level OK
            }
            else {
                $script:Stats.Skipped++
            }
        }
        catch {
            $script:Stats.Errors++
            Write-Log "  FAILED: $_" -Level ERROR
        }

        Start-Sleep -Milliseconds $ThrottleDelayMs
    }

    Save-State -State $state

    $elapsed = (Get-Date) - $script:Stats.Start
    Write-Host ""
    Write-Host "  Done in $($elapsed.ToString('mm\:ss'))" -ForegroundColor Cyan
    Write-Host "  Processed $($script:Stats.Processed) | Deployed $($script:Stats.Deployed) | Skipped $($script:Stats.Skipped) | Errors $($script:Stats.Errors)" `
        -ForegroundColor $(if($script:Stats.Errors -gt 0){'Yellow'}else{'Green'})
    Write-Host "  Log: $LogPath" -ForegroundColor DarkGray
    Write-Host ""
}

Main
