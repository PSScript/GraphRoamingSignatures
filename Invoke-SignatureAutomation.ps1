<#
.SYNOPSIS
    Invoke-SignatureAutomation.ps1 — Delta-aware bulk signature deployment
    Multi-company template assignment with VIP overrides and checksum change tracking.

.DESCRIPTION
    Designed for scheduled/unattended execution (Azure Automation, Task Scheduler, cron).

    Assignment hierarchy (first match wins):
      1. User-level override  (VIP map — CEO wants his LinkedIn QR code)
      2. Security group match (SG-Signature-TemplateX)
      3. companyName match    (Entra ID companyName → template mapping)
      4. Default template     (fallback)

    Change detection:
      - Graph delta query: /users/delta — only process users whose properties changed
      - Checksum in extensionAttribute15: SHA256(user props + template hash)
        → Skip deployment if hash matches (nothing changed for this user)
      - Template file hash change forces full re-run (rebrand scenario)

    State persistence:
      - Delta token + template hashes stored in state.json (local or Azure Blob)
      - extensionAttribute15 on each user: "SIG:<hash>|<timestamp>"

.PARAMETER ConfigPath
    Path to the YAML/JSON company-template mapping config.

.PARAMETER StatePath
    Path to persist delta tokens and template hashes between runs.

.PARAMETER Force
    Ignore checksums, redeploy ALL users (use after template redesign).

.PARAMETER WhatIf
    Preview mode — show what would change without deploying.

.EXAMPLE
    # Normal scheduled run (only changed users)
    .\Invoke-SignatureAutomation.ps1 -ConfigPath .\config.json -StatePath .\state.json

    # Force full redeploy after rebrand
    .\Invoke-SignatureAutomation.ps1 -ConfigPath .\config.json -StatePath .\state.json -Force

    # Preview what would change
    .\Invoke-SignatureAutomation.ps1 -ConfigPath .\config.json -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [string]$StatePath = '.\signature-state.json',

    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [Parameter(Mandatory)]
    [string]$ClientSecret,

    [switch]$Force,

    [string]$ChecksumAttribute = 'extensionAttribute15',

    [string]$LogPath = ".\signature-auto-$(Get-Date -Format 'yyyyMMdd-HHmmss').log",

    [int]$ThrottleDelayMs = 100,

    [int]$BatchSize = 20
)

$ErrorActionPreference = 'Stop'

#region ═══════════════════════════════════════════════════════════════════
# LOGGING
#endregion ════════════════════════════════════════════════════════════════

$script:Stats = @{ Processed=0; Deployed=0; Skipped=0; Errors=0; StartTime=Get-Date }

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    switch ($Level) {
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'DEBUG'   { Write-Verbose $line }
        default   { Write-Host $line }
    }
    $line | Out-File -FilePath $LogPath -Append -Encoding utf8
}

#region ═══════════════════════════════════════════════════════════════════
# CONFIGURATION
#endregion ════════════════════════════════════════════════════════════════

<#
    Config format (config.json):
    {
        "signatureNamePrefix": "Corporate",
        "checksumAttribute": "extensionAttribute15",
        "checksumPrefix": "SIG:",

        "companies": {
            "Contoso Group": {
                "template": "./templates/contoso.htm",
                "signatureName": "Contoso Signatur"
            },
            "Contoso Cloud": {
                "template": "./templates/contoso-bit.htm",
                "signatureName": "Contoso Cloud Signatur"
            },
            "Fabrikam GmbH": {
                "template": "./templates/fabrikam.htm",
                "signatureName": "Fabrikam Signatur"
            }
        },

        "groupOverrides": {
            "SG-Signature-Marketing": {
                "template": "./templates/marketing-campaign.htm",
                "signatureName": "Marketing Campaign Q2",
                "groupId": "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
            }
        },

        "vipOverrides": {
            "ceo@contoso.de": {
                "template": "./templates/vip-ceo-contoso.htm",
                "signatureName": "CEO Signatur"
            },
            "md.north@contoso-bit.de": {
                "template": "./templates/vip-md-north.htm",
                "signatureName": "MD North"
            },
            "md.south@contoso-bit.de": {
                "template": "./templates/vip-md-south.htm",
                "signatureName": "MD South"
            }
        },

        "defaultTemplate": {
            "template": "./templates/default.htm",
            "signatureName": "Standard Signatur"
        },

        "excludeUPNs": [
            "admin@contoso.com",
            "noreply@contoso.com"
        ],

        "excludeCompanyNames": [
            "Service Accounts"
        ]
    }
#>

function Load-Config {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Config file not found: $Path"
    }

    $config = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json

    # Validate required fields
    if (-not $config.companies -and -not $config.defaultTemplate) {
        throw "Config must define at least 'companies' or 'defaultTemplate'"
    }

    # Preload and hash all templates
    $templateCache = @{}
    $allTemplateRefs = @()

    # Collect all template paths
    if ($config.companies) {
        foreach ($prop in $config.companies.PSObject.Properties) {
            $allTemplateRefs += @{ Key = "company:$($prop.Name)"; Config = $prop.Value }
        }
    }
    if ($config.groupOverrides) {
        foreach ($prop in $config.groupOverrides.PSObject.Properties) {
            $allTemplateRefs += @{ Key = "group:$($prop.Name)"; Config = $prop.Value }
        }
    }
    if ($config.vipOverrides) {
        foreach ($prop in $config.vipOverrides.PSObject.Properties) {
            $allTemplateRefs += @{ Key = "vip:$($prop.Name)"; Config = $prop.Value }
        }
    }
    if ($config.defaultTemplate) {
        $allTemplateRefs += @{ Key = "default"; Config = $config.defaultTemplate }
    }

    foreach ($ref in $allTemplateRefs) {
        $tplPath = $ref.Config.template
        if (-not $tplPath) { continue }

        # Resolve relative paths
        if (-not [System.IO.Path]::IsPathRooted($tplPath)) {
            $tplPath = Join-Path (Split-Path $Path -Parent) $tplPath
        }

        if (-not (Test-Path -LiteralPath $tplPath)) {
            Write-Log "Template not found: $tplPath (referenced by $($ref.Key))" -Level ERROR
            continue
        }

        if (-not $templateCache.ContainsKey($tplPath)) {
            $content = Read-HtmlTemplate -Path $tplPath
            $hash = Get-StringHash -Text $content
            $templateCache[$tplPath] = @{
                Content = $content
                Hash    = $hash
                Path    = $tplPath
            }
            Write-Log "  Template loaded: $tplPath (hash: $($hash.Substring(0,8)))"
        }
    }

    return @{
        Config   = $config
        Cache    = $templateCache
    }
}

#region ═══════════════════════════════════════════════════════════════════
# ENCODING + HASHING
#endregion ════════════════════════════════════════════════════════════════

function Read-HtmlTemplate {
    param([string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $encoding = Detect-Encoding -Bytes $bytes
    $content = [System.IO.File]::ReadAllText($Path, $encoding)
    if ($content.Length -gt 0 -and $content[0] -eq [char]0xFEFF) { $content = $content.Substring(1) }
    $content = $content -replace "`r?`n", "`r`n"
    # HTML-entity encode Windows-1252 specials for email client compat
    $content = $content -replace [char]0x201C, '&ldquo;' -replace [char]0x201D, '&rdquo;'
    $content = $content -replace [char]0x2018, '&lsquo;' -replace [char]0x2019, '&rsquo;'
    $content = $content -replace [char]0x2013, '&ndash;'  -replace [char]0x2014, '&mdash;'
    $content = $content -replace [char]0x20AC, '&euro;'   -replace [char]0x2122, '&trade;'
    if ($content -match 'charset\s*=\s*["\x27]?\s*[\w-]+') {
        $content = $content -replace '(charset\s*=\s*["\x27]?\s*)[\w-]+', '${1}utf-8'
    }
    return $content
}

function Detect-Encoding {
    param([byte[]]$Bytes)
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8
    }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) {
        return [System.Text.Encoding]::Unicode
    }
    $headBytes = if ($Bytes.Length -gt 4096) { $Bytes[0..4095] } else { $Bytes }
    $head = [System.Text.Encoding]::ASCII.GetString($headBytes)
    if ($head -match 'charset\s*=\s*["\x27]?\s*([\w-]+)') {
        $cs = $Matches[1].ToLower()
        if ($cs -match 'windows.?1252')  { return [System.Text.Encoding]::GetEncoding(1252) }
        if ($cs -match 'iso.?8859.?1')   { return [System.Text.Encoding]::GetEncoding(28591) }
    }
    for ($i = 0; $i -lt [Math]::Min($Bytes.Length, 8192); $i++) {
        if ($Bytes[$i] -ge 0x80 -and $Bytes[$i] -le 0x9F -and $Bytes[$i] -notin @(0x81,0x8D,0x8F,0x90,0x9D)) {
            return [System.Text.Encoding]::GetEncoding(1252)
        }
    }
    return [System.Text.Encoding]::UTF8
}

function Get-StringHash {
    param([string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $hash = $sha.ComputeHash($bytes)
    return [BitConverter]::ToString($hash).Replace('-', '').ToLower()
}

#region ═══════════════════════════════════════════════════════════════════
# AUTHENTICATION
#endregion ════════════════════════════════════════════════════════════════

function Get-GraphToken {
    $body = @{
        grant_type    = 'client_credentials'
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
    }
    $response = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/${TenantId}/oauth2/v2.0/token" `
        -ContentType 'application/x-www-form-urlencoded' -Body $body
    $script:Token = $response.access_token
    $script:TokenExpires = (Get-Date).AddSeconds($response.expires_in - 300)
}

function Invoke-Graph {
    param([string]$Method='GET', [string]$Uri, [object]$Body, [switch]$Beta, [int]$Retries=3)

    if (-not $script:Token -or (Get-Date) -ge $script:TokenExpires) { Get-GraphToken }

    $base = if ($Beta) { 'https://graph.microsoft.com/beta' } else { 'https://graph.microsoft.com/v1.0' }
    $fullUri = if ($Uri.StartsWith('http')) { $Uri } else { "${base}${Uri}" }

    $headers = @{
        Authorization      = "Bearer $($script:Token)"
        'Content-Type'     = 'application/json'
        ConsistencyLevel   = 'eventual'
    }

    for ($r = 0; $r -lt $Retries; $r++) {
        try {
            $params = @{ Method=$Method; Uri=$fullUri; Headers=$headers }
            if ($Body -and $Method -ne 'GET') {
                $params.Body = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 -Compress }
            }
            return Invoke-RestMethod @params -ErrorAction Stop
        }
        catch {
            $code = $_.Exception.Response.StatusCode.Value__
            if ($code -eq 429) {
                $wait = 30
                try { $wait = [int]($_.Exception.Response.Headers | Where-Object Key -eq 'Retry-After').Value[0] } catch {}
                Write-Log "  Throttled. Waiting ${wait}s (retry $($r+1)/$Retries)" -Level WARN
                Start-Sleep -Seconds $wait
            }
            elseif ($code -in 500,502,503,504 -and $r -lt $Retries-1) {
                Start-Sleep -Seconds ([math]::Pow(2,$r)*3)
            }
            else { throw }
        }
    }
}

#region ═══════════════════════════════════════════════════════════════════
# DELTA QUERY — only fetch users whose properties changed
#endregion ════════════════════════════════════════════════════════════════

function Get-ChangedUsers {
    <#
    .SYNOPSIS
        Uses Graph delta query to fetch only users whose signature-relevant properties
        changed since the last run. Falls back to full fetch on first run or if delta
        token expired (>30 days).
    #>
    param([hashtable]$State)

    $select = @(
        'id','userPrincipalName','displayName','givenName','surname','mail',
        'jobTitle','department','companyName','officeLocation',
        'businessPhones','mobilePhone','faxNumber',
        'streetAddress','city','state','postalCode','country',
        'onPremisesExtensionAttributes','mailNickname','memberOf'
    ) -join ','

    $deltaLink = $State.deltaLink
    $users = [System.Collections.Generic.List[object]]::new()
    $isFullSync = $false

    if ($deltaLink -and -not $Force) {
        Write-Log "Using delta token from last run"
        try {
            $response = Invoke-Graph -Uri $deltaLink
        }
        catch {
            Write-Log "Delta token expired or invalid — falling back to full sync" -Level WARN
            $deltaLink = $null
        }
    }

    if (-not $deltaLink -or $Force) {
        $isFullSync = $true
        Write-Log "$(if($Force){'Forced full sync'}else{'First run — full sync'})"
        $response = Invoke-Graph -Uri "/users/delta?`$select=${select}&`$top=200" -Beta
    }

    # Page through results
    while ($response) {
        if ($response.value) {
            $users.AddRange($response.value)
        }

        # Capture delta link for next run
        if ($response.'@odata.deltaLink') {
            $State.deltaLink = $response.'@odata.deltaLink'
            $response = $null
        }
        elseif ($response.'@odata.nextLink') {
            Start-Sleep -Milliseconds $ThrottleDelayMs
            $response = Invoke-Graph -Uri $response.'@odata.nextLink'
        }
        else {
            $response = $null
        }
    }

    Write-Log "Delta query returned $($users.Count) users (fullSync: $isFullSync)"
    return @{ Users = $users; IsFullSync = $isFullSync }
}

#region ═══════════════════════════════════════════════════════════════════
# TEMPLATE ASSIGNMENT — hierarchy resolution
#endregion ════════════════════════════════════════════════════════════════

function Resolve-UserTemplate {
    <#
    .SYNOPSIS
        Determines which template a user gets based on the assignment hierarchy:
        1. VIP override (by UPN)
        2. Security group override (by group membership)
        3. companyName match
        4. Default template
    #>
    param(
        [psobject]$User,
        [hashtable]$ConfigData,
        [hashtable]$GroupMembershipCache
    )

    $config = $ConfigData.Config
    $cache  = $ConfigData.Cache
    $upn    = $User.userPrincipalName

    # ─── Priority 1: VIP override (exact UPN match) ─────────────────
    if ($config.vipOverrides) {
        $vipMatch = $config.vipOverrides.PSObject.Properties | Where-Object { $_.Name -ieq $upn }
        if ($vipMatch) {
            $tplPath = Resolve-TemplatePath -RelPath $vipMatch.Value.template -ConfigPath $ConfigPath
            if ($cache.ContainsKey($tplPath)) {
                Write-Log "    Assignment: VIP override ($upn)" -Level DEBUG
                return @{
                    Source        = "VIP:$upn"
                    SignatureName = $vipMatch.Value.signatureName
                    Template      = $cache[$tplPath]
                }
            }
        }
    }

    # ─── Priority 2: Security group override ─────────────────────────
    if ($config.groupOverrides) {
        foreach ($prop in $config.groupOverrides.PSObject.Properties) {
            $groupConfig = $prop.Value
            $groupId = $groupConfig.groupId
            if (-not $groupId) { continue }

            # Check if user is member (use cache to avoid per-user API calls)
            if (-not $GroupMembershipCache.ContainsKey($groupId)) {
                Write-Log "    Fetching members for group: $($prop.Name) ($groupId)" -Level DEBUG
                try {
                    $members = @()
                    $resp = Invoke-Graph -Uri "/groups/${groupId}/members?`$select=id&`$top=999" -Beta
                    while ($resp) {
                        $members += $resp.value.id
                        if ($resp.'@odata.nextLink') {
                            $resp = Invoke-Graph -Uri $resp.'@odata.nextLink'
                        } else { $resp = $null }
                    }
                    $GroupMembershipCache[$groupId] = $members
                }
                catch {
                    Write-Log "    Failed to fetch group $groupId members: $_" -Level WARN
                    $GroupMembershipCache[$groupId] = @()
                }
            }

            if ($User.id -in $GroupMembershipCache[$groupId]) {
                $tplPath = Resolve-TemplatePath -RelPath $groupConfig.template -ConfigPath $ConfigPath
                if ($cache.ContainsKey($tplPath)) {
                    Write-Log "    Assignment: Group override ($($prop.Name))" -Level DEBUG
                    return @{
                        Source        = "Group:$($prop.Name)"
                        SignatureName = $groupConfig.signatureName
                        Template      = $cache[$tplPath]
                    }
                }
            }
        }
    }

    # ─── Priority 3: companyName match ───────────────────────────────
    if ($config.companies -and $User.companyName) {
        $companyMatch = $config.companies.PSObject.Properties | Where-Object { $_.Name -ieq $User.companyName }
        if ($companyMatch) {
            $tplPath = Resolve-TemplatePath -RelPath $companyMatch.Value.template -ConfigPath $ConfigPath
            if ($cache.ContainsKey($tplPath)) {
                Write-Log "    Assignment: companyName '$($User.companyName)'" -Level DEBUG
                return @{
                    Source        = "Company:$($User.companyName)"
                    SignatureName = $companyMatch.Value.signatureName
                    Template      = $cache[$tplPath]
                }
            }
        }
    }

    # ─── Priority 4: Default template ────────────────────────────────
    if ($config.defaultTemplate) {
        $tplPath = Resolve-TemplatePath -RelPath $config.defaultTemplate.template -ConfigPath $ConfigPath
        if ($cache.ContainsKey($tplPath)) {
            Write-Log "    Assignment: Default template" -Level DEBUG
            return @{
                Source        = "Default"
                SignatureName = $config.defaultTemplate.signatureName
                Template      = $cache[$tplPath]
            }
        }
    }

    Write-Log "    Assignment: NO TEMPLATE FOUND" -Level WARN
    return $null
}

function Resolve-TemplatePath {
    param([string]$RelPath, [string]$ConfigPath)
    if ([System.IO.Path]::IsPathRooted($RelPath)) { return $RelPath }
    return Join-Path (Split-Path $ConfigPath -Parent) $RelPath
}

#region ═══════════════════════════════════════════════════════════════════
# CHECKSUM — change detection via extensionAttribute
#endregion ════════════════════════════════════════════════════════════════

function Get-SignatureChecksum {
    <#
    .SYNOPSIS
        Computes a SHA256 hash over the user's signature-relevant properties
        + the template hash. If this matches what's stored in extensionAttribute,
        the user's signature hasn't changed and we can skip deployment.
    #>
    param(
        [psobject]$User,
        [string]$TemplateHash
    )

    # Concatenate all properties that affect the signature output
    $sigString = @(
        $User.displayName
        $User.givenName
        $User.surname
        $User.mail
        $User.jobTitle
        $User.department
        $User.companyName
        $User.officeLocation
        ($User.businessPhones -join '|')
        $User.mobilePhone
        $User.faxNumber
        $User.streetAddress
        $User.city
        $User.state
        $User.postalCode
        $User.country
        $TemplateHash  # Template change = checksum change = redeploy
    ) -join '|'

    $hash = Get-StringHash -Text $sigString
    return $hash.Substring(0, 12)  # First 12 chars is plenty for change detection
}

function Get-StoredChecksum {
    <#
    .SYNOPSIS
        Reads the checksum from the user's extensionAttribute.
        Format: "SIG:<hash>|<ISO8601 timestamp>"
    #>
    param([psobject]$User, [string]$Prefix = 'SIG:')

    $attrValue = $null
    $attrMap = @{
        'extensionAttribute15' = 'onPremisesExtensionAttributes.extensionAttribute15'
        'extensionAttribute14' = 'onPremisesExtensionAttributes.extensionAttribute14'
        'extensionAttribute13' = 'onPremisesExtensionAttributes.extensionAttribute13'
    }

    if ($User.onPremisesExtensionAttributes) {
        $attrValue = $User.onPremisesExtensionAttributes.$($ChecksumAttribute -replace 'extensionAttribute','extensionAttribute')
    }

    if ($attrValue -and $attrValue.StartsWith($Prefix)) {
        $parts = $attrValue.Substring($Prefix.Length).Split('|')
        return @{
            Hash      = $parts[0]
            Timestamp = if ($parts.Length -gt 1) { $parts[1] } else { $null }
        }
    }

    return $null
}

function Set-StoredChecksum {
    <#
    .SYNOPSIS
        Writes the checksum to the user's extensionAttribute via Graph.
        Only works for cloud-only attributes or if writeback is configured.

        For hybrid environments with AD Connect, extensionAttributes sync FROM on-prem.
        In that case, you need to either:
        a) Write to an attribute that ISN'T synced (customSecurityAttributes or a free extAttr)
        b) Write back to on-prem AD via LDAP
        c) Use a SharePoint list or Azure Table instead
    #>
    param(
        [string]$UserId,
        [string]$Hash,
        [string]$Prefix = 'SIG:'
    )

    $value = "${Prefix}${Hash}|$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ')"

    # Map attribute name to Graph property path
    $attrName = $ChecksumAttribute
    $body = @{
        onPremisesExtensionAttributes = @{
            $attrName = $value
        }
    }

    try {
        Invoke-Graph -Method PATCH -Uri "/users/${UserId}" -Body $body -Beta | Out-Null
        return $true
    }
    catch {
        $errMsg = "$_"
        if ($errMsg -match 'cloud-mastered' -or $errMsg -match 'DirSync') {
            Write-Log "      Cannot write $attrName — attribute is AD-synced. Use a cloud-only attribute or SharePoint list." -Level WARN
            return $false
        }
        Write-Log "      Failed to write checksum: $errMsg" -Level WARN
        return $false
    }
}

#region ═══════════════════════════════════════════════════════════════════
# VARIABLE EXPANSION
#endregion ════════════════════════════════════════════════════════════════

function Expand-TemplateVariables {
    param([string]$Template, [psobject]$User, [psobject]$Manager)

    $r = $Template
    $map = @{
        '{{DisplayName}}'  = $User.displayName
        '{{GivenName}}'    = $User.givenName
        '{{Surname}}'      = $User.surname
        '{{Mail}}'         = $User.mail
        '{{UPN}}'          = $User.userPrincipalName
        '{{JobTitle}}'     = $User.jobTitle
        '{{Department}}'   = $User.department
        '{{Company}}'      = $User.companyName
        '{{Office}}'       = $User.officeLocation
        '{{Phone}}'        = if ($User.businessPhones) { $User.businessPhones[0] } else { '' }
        '{{Phone2}}'       = if ($User.businessPhones -and $User.businessPhones.Count -gt 1) { $User.businessPhones[1] } else { '' }
        '{{Mobile}}'       = $User.mobilePhone
        '{{Fax}}'          = $User.faxNumber
        '{{Street}}'       = $User.streetAddress
        '{{City}}'         = $User.city
        '{{State}}'        = $User.state
        '{{PostalCode}}'   = $User.postalCode
        '{{Country}}'      = $User.country
        '{{MailNickname}}'  = $User.mailNickname
        '{{ManagerName}}'  = if ($Manager) { $Manager.displayName } else { '' }
        '{{ManagerMail}}'  = if ($Manager) { $Manager.mail } else { '' }
        '{{ManagerTitle}}' = if ($Manager) { $Manager.jobTitle } else { '' }
    }

    for ($i = 1; $i -le 15; $i++) {
        $val = ''
        if ($User.onPremisesExtensionAttributes) {
            $val = $User.onPremisesExtensionAttributes."extensionAttribute$i"
        }
        $map["{{ExtAttr$i}}"] = $val
    }

    foreach ($key in $map.Keys) {
        $val = if ($null -eq $map[$key]) { '' } else { [System.Web.HttpUtility]::HtmlEncode($map[$key]) }
        $r = $r.Replace($key, $val)
    }

    # Clean up unreplaced variables
    $r = $r -replace '\{\{\w+\}\}', ''
    # Remove empty rows
    $r = $r -replace '<tr>\s*(<td[^>]*>\s*(<[^>]*>\s*)*\s*(</[^>]*>\s*)*\s*</td>\s*)+</tr>', ''
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

#region ═══════════════════════════════════════════════════════════════════
# DEPLOYMENT — Graph Beta UserConfiguration API
#endregion ════════════════════════════════════════════════════════════════

function Deploy-Signature {
    param(
        [string]$UserId,
        [string]$UPN,
        [string]$SignatureHtml,
        [string]$SignatureText,
        [string]$SignatureName
    )

    # Get Inbox folder ID
    # inbox is a well-known folder name — no Mail.ReadWrite needed

    # Try UserConfiguration API (Graph Beta)
    $configName = 'OWA.UserOptions'

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

    $base64Dict = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($dictXml))

    $body = @{
        '@odata.type' = '#microsoft.graph.userConfiguration'
        structuredData = $base64Dict
    }

    Invoke-Graph -Method PATCH `
        -Uri "/users/${UserId}/mailFolders/inbox/userConfigurations/${configName}" `
        -Body $body -Beta | Out-Null
}

#region ═══════════════════════════════════════════════════════════════════
# STATE PERSISTENCE
#endregion ════════════════════════════════════════════════════════════════

function Load-State {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        return Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable
    }
    return @{ deltaLink = $null; templateHashes = @{}; lastRun = $null }
}

function Save-State {
    param([hashtable]$State, [string]$Path)
    $State.lastRun = Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ'
    $State | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding utf8
}

#region ═══════════════════════════════════════════════════════════════════
# MAIN ORCHESTRATOR
#endregion ════════════════════════════════════════════════════════════════

function Main {
    Write-Host ""
    Write-Host "  Signature Automation — Delta-Aware Bulk Deployment" -ForegroundColor Cyan
    Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor DarkGray
    Write-Host ""

    # Load config and templates
    Write-Log "Loading configuration: $ConfigPath"
    $configData = Load-Config -Path $ConfigPath

    # Load state (delta token, template hashes)
    $state = Load-State -Path $StatePath
    Write-Log "State loaded (lastRun: $($state.lastRun ?? 'never'))"

    # Check if any template changed since last run (forces full redeploy)
    $templateChanged = $false
    foreach ($tplPath in $configData.Cache.Keys) {
        $currentHash = $configData.Cache[$tplPath].Hash
        $storedHash = $state.templateHashes[$tplPath]
        if ($storedHash -ne $currentHash) {
            Write-Log "  Template changed: $tplPath (was: $($storedHash?.Substring(0,8) ?? 'new'), now: $($currentHash.Substring(0,8)))" -Level WARN
            $templateChanged = $true
        }
        $state.templateHashes[$tplPath] = $currentHash
    }

    if ($templateChanged -and -not $Force) {
        Write-Log "Template(s) changed — forcing full sync for all users" -Level WARN
        $Force = $true
    }

    # Authenticate
    Write-Log "Authenticating..."
    Get-GraphToken
    Write-Log "  Token acquired" -Level SUCCESS

    # Fetch changed users via delta query
    $deltaResult = Get-ChangedUsers -State $state
    $changedUsers = $deltaResult.Users

    # Filter out excluded users
    $config = $configData.Config
    $excludeUPNs = @()
    if ($config.excludeUPNs) { $excludeUPNs = $config.excludeUPNs }
    $excludeCompanies = @()
    if ($config.excludeCompanyNames) { $excludeCompanies = $config.excludeCompanyNames }

    $eligibleUsers = $changedUsers | Where-Object {
        $_.mail -and
        $_.userPrincipalName -and
        $_.userPrincipalName -notin $excludeUPNs -and
        $_.companyName -notin $excludeCompanies -and
        -not ($_.userPrincipalName -match '^(?:admin|noreply|shared|room|equipment)')
    }

    Write-Log "Eligible users after filtering: $($eligibleUsers.Count) (from $($changedUsers.Count) delta results)"

    # Group membership cache (lazy-loaded per group on first access)
    $groupCache = @{}

    # Manager cache
    $managerCache = @{}

    # Process users
    $total = @($eligibleUsers).Count
    $i = 0

    foreach ($user in $eligibleUsers) {
        $i++
        $upn = $user.userPrincipalName
        $pct = if ($total -gt 0) { [math]::Round($i / $total * 100) } else { 0 }
        $script:Stats.Processed++

        Write-Log "[$i/$total] (${pct}%) $upn — $($user.companyName ?? '(no company)')"

        try {
            # Resolve template assignment
            $assignment = Resolve-UserTemplate -User $user -ConfigData $configData -GroupMembershipCache $groupCache

            if (-not $assignment) {
                Write-Log "    SKIP: No template matched" -Level WARN
                $script:Stats.Skipped++
                continue
            }

            Write-Log "    Template: $($assignment.Source) → $($assignment.SignatureName)"

            # Checksum comparison — skip if nothing changed
            $expectedChecksum = Get-SignatureChecksum -User $user -TemplateHash $assignment.Template.Hash
            $storedChecksum = Get-StoredChecksum -User $user

            if (-not $Force -and $storedChecksum -and $storedChecksum.Hash -eq $expectedChecksum) {
                Write-Log "    SKIP: Checksum match ($expectedChecksum) — no changes" -Level DEBUG
                $script:Stats.Skipped++
                continue
            }

            Write-Log "    Checksum: stored=$($storedChecksum?.Hash ?? 'none') expected=$expectedChecksum — DEPLOYING"

            # Fetch manager (cached)
            $manager = $null
            if (-not $managerCache.ContainsKey($user.id)) {
                try {
                    $manager = Invoke-Graph -Uri "/users/$($user.id)/manager?`$select=displayName,mail,jobTitle" -Beta
                    $managerCache[$user.id] = $manager
                } catch {
                    $managerCache[$user.id] = $null
                }
            }
            $manager = $managerCache[$user.id]

            # Expand template
            $sigHtml = Expand-TemplateVariables -Template $assignment.Template.Content -User $user -Manager $manager
            $sigText = ConvertTo-PlainText -Html $sigHtml

            # Deploy
            if ($PSCmdlet.ShouldProcess($upn, "Deploy signature '$($assignment.SignatureName)' via $($assignment.Source)")) {
                Deploy-Signature -UserId $user.id -UPN $upn `
                    -SignatureHtml $sigHtml -SignatureText $sigText `
                    -SignatureName $assignment.SignatureName

                # Write checksum back
                $null = Set-StoredChecksum -UserId $user.id -Hash $expectedChecksum

                $script:Stats.Deployed++
                Write-Log "    DEPLOYED ($($assignment.Source))" -Level SUCCESS
            }
            else {
                Write-Log "    [WhatIf] Would deploy $($assignment.SignatureName)" -Level WARN
                $script:Stats.Skipped++
            }
        }
        catch {
            $script:Stats.Errors++
            Write-Log "    ERROR: $_" -Level ERROR
        }

        # Throttle
        Start-Sleep -Milliseconds $ThrottleDelayMs
    }

    # Save state for next run
    Save-State -State $state -Path $StatePath

    # Summary
    $elapsed = (Get-Date) - $script:Stats.StartTime
    Write-Host ""
    Write-Host "  ══════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  COMPLETE  $($elapsed.ToString('mm\:ss')) elapsed" -ForegroundColor Cyan
    Write-Host "  Processed : $($script:Stats.Processed)" -ForegroundColor White
    Write-Host "  Deployed  : $($script:Stats.Deployed)" -ForegroundColor Green
    Write-Host "  Skipped   : $($script:Stats.Skipped)" -ForegroundColor DarkGray
    Write-Host "  Errors    : $($script:Stats.Errors)" -ForegroundColor $(if($script:Stats.Errors -gt 0){'Red'}else{'DarkGray'})
    Write-Host "  State     : $StatePath" -ForegroundColor DarkGray
    Write-Host "  Log       : $LogPath" -ForegroundColor DarkGray
    Write-Host "  ══════════════════════════════════════════════════" -ForegroundColor Cyan
}

Add-Type -AssemblyName System.Web
Main
