<#
.SYNOPSIS
    Manage-OutlookSignatures.ps1 — Enterprise Signature Management via Microsoft Graph
    Supports roaming signatures (Graph Beta UserConfiguration API), OWA classic, and transport rules.

.DESCRIPTION
    Three deployment strategies:
      1. Graph Beta UserConfiguration API (MailboxConfigItem.ReadWrite) — roaming signatures
      2. Graph mailboxSettings API — OOF and classic OWA settings
      3. EXO Transport Rules — server-side injection (all clients)

    Requires Entra ID App Registration with:
      - MailboxConfigItem.ReadWrite (delegated or app)
      # Mail.ReadWrite NOT needed — inbox is a well-known folder name
      - User.Read.All (delegated or app)
      - MailboxSettings.ReadWrite (delegated or app)

.NOTES
    Author : Jan Hübener / Contoso Group
    Date   : 2026-03-23
    Ref    : Graph Beta userConfiguration API (Jan 2026)
             https://devblogs.microsoft.com/microsoft365dev/introducing-the-microsoft-graph-user-configuration-api-preview/
    License: MIT

.PARAMETER TenantId
    Azure AD / Entra ID Tenant ID

.PARAMETER ClientId
    App Registration Client ID

.PARAMETER ClientSecret
    App Registration Client Secret (for app-only flow). Omit for delegated/interactive.

.PARAMETER TemplatePath
    Path to HTML signature template file(s). Supports Windows-1252, ISO-8859-1, UTF-8, UTF-16.

.PARAMETER UserUPN
    Target user UPN. Use '*' or omit for all licensed users.

.PARAMETER Strategy
    Deployment strategy: 'Roaming', 'OWA', 'TransportRule', or 'All'

.PARAMETER DryRun
    Preview changes without applying.

.EXAMPLE
    # Interactive / delegated flow — single user roaming signature
    .\Manage-OutlookSignatures.ps1 -TenantId "abc-123" -ClientId "def-456" `
        -TemplatePath ".\templates\corporate.htm" -UserUPN "user@contoso.com" -Strategy Roaming

.EXAMPLE
    # App-only flow — all users
    .\Manage-OutlookSignatures.ps1 -TenantId "abc-123" -ClientId "def-456" `
        -ClientSecret "s3cret" -TemplatePath ".\templates\corporate.htm" -Strategy All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [string]$ClientSecret,

    [Parameter(Mandatory)]
    [string]$TemplatePath,

    [string]$UserUPN = '*',

    [ValidateSet('Roaming', 'OWA', 'TransportRule', 'All')]
    [string]$Strategy = 'Roaming',

    [string]$SignatureName = 'Corporate Signature',

    [switch]$SetAsDefault,

    [switch]$SetForReply,

    [switch]$DryRun,

    [string]$LogPath = ".\signature-deploy-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
)

$ErrorActionPreference = 'Stop'
$script:GraphBaseUri = 'https://graph.microsoft.com'
$script:GraphBeta    = "$($script:GraphBaseUri)/beta"
$script:GraphV1      = "$($script:GraphBaseUri)/v1.0"

#region ═══════════════════════════════════════════════════════════════════
# LOGGING
#endregion ════════════════════════════════════════════════════════════════
function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','SUCCESS')]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    switch ($Level) {
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        default   { Write-Host $line }
    }
    $line | Out-File -FilePath $LogPath -Append -Encoding utf8
}

#region ═══════════════════════════════════════════════════════════════════
# ENCODING ENGINE — Windows-1252 / ISO-8859-1 / UTF-8 / UTF-16 aware
#endregion ════════════════════════════════════════════════════════════════

function Get-HtmlFileEncoding {
    <#
    .SYNOPSIS
        Detects encoding of an HTML file via BOM, charset meta tag, or byte heuristics.
        Returns [System.Text.Encoding] object.
    #>
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)

    # BOM detection
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        Write-Log "  Encoding: UTF-8 (BOM detected)" -Level INFO
        return [System.Text.Encoding]::UTF8
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        Write-Log "  Encoding: UTF-16 LE (BOM detected)" -Level INFO
        return [System.Text.Encoding]::Unicode
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        Write-Log "  Encoding: UTF-16 BE (BOM detected)" -Level INFO
        return [System.Text.Encoding]::BigEndianUnicode
    }

    # Parse charset from meta tag (first 4KB)
    $headBytes = if ($bytes.Length -gt 4096) { $bytes[0..4095] } else { $bytes }
    $asciiHead = [System.Text.Encoding]::ASCII.GetString($headBytes)

    if ($asciiHead -match 'charset\s*=\s*["\''']?\s*([\w-]+)') {
        $charset = $Matches[1].ToLower()
        Write-Log "  Encoding: charset meta tag found: $charset" -Level INFO

        switch -Regex ($charset) {
            'windows.?1252'    { return [System.Text.Encoding]::GetEncoding(1252) }
            'iso.?8859.?1'     { return [System.Text.Encoding]::GetEncoding(28591) }
            'iso.?8859.?15'    { return [System.Text.Encoding]::GetEncoding(28605) }
            'utf.?8'           { return [System.Text.Encoding]::UTF8 }
            'utf.?16'          { return [System.Text.Encoding]::Unicode }
            'ascii'            { return [System.Text.Encoding]::ASCII }
            default {
                try {
                    return [System.Text.Encoding]::GetEncoding($charset)
                } catch {
                    Write-Log "  Unknown charset '$charset', falling back to heuristic" -Level WARN
                }
            }
        }
    }

    # Heuristic: scan for Windows-1252 specific bytes (0x80-0x9F range)
    # These bytes are undefined in ISO-8859-1 but map to specific chars in Win-1252
    $hasWin1252Specifics = $false
    $hasHighBytes = $false
    $invalidUtf8Sequences = 0
    $scanLength = [Math]::Min($bytes.Length, 8192)

    for ($i = 0; $i -lt $scanLength; $i++) {
        $b = $bytes[$i]
        if ($b -ge 0x80 -and $b -le 0x9F) {
            # These specific bytes only make sense in Windows-1252
            # 0x81, 0x8D, 0x8F, 0x90, 0x9D are undefined even in Win-1252
            if ($b -notin @(0x81, 0x8D, 0x8F, 0x90, 0x9D)) {
                $hasWin1252Specifics = $true
            }
        }
        if ($b -gt 0x7F) { $hasHighBytes = $true }

        # UTF-8 multi-byte validation
        if ($b -ge 0xC0 -and $b -le 0xDF) {
            # Expect 1 continuation byte
            if (($i + 1) -ge $bytes.Length -or ($bytes[$i + 1] -band 0xC0) -ne 0x80) {
                $invalidUtf8Sequences++
            }
        } elseif ($b -ge 0xE0 -and $b -le 0xEF) {
            # Expect 2 continuation bytes
            if (($i + 2) -ge $bytes.Length -or
                ($bytes[$i + 1] -band 0xC0) -ne 0x80 -or
                ($bytes[$i + 2] -band 0xC0) -ne 0x80) {
                $invalidUtf8Sequences++
            }
        }
    }

    if ($hasWin1252Specifics) {
        Write-Log "  Encoding: Windows-1252 (heuristic — found 0x80-0x9F range bytes)" -Level INFO
        return [System.Text.Encoding]::GetEncoding(1252)
    }

    if ($hasHighBytes -and $invalidUtf8Sequences -gt 0) {
        Write-Log "  Encoding: ISO-8859-1 (heuristic — high bytes but invalid UTF-8 sequences: $invalidUtf8Sequences)" -Level INFO
        return [System.Text.Encoding]::GetEncoding(28591)
    }

    Write-Log "  Encoding: UTF-8 (default)" -Level INFO
    return [System.Text.Encoding]::UTF8
}

function Read-HtmlTemplate {
    <#
    .SYNOPSIS
        Reads an HTML file with automatic encoding detection and normalizes to UTF-8.
        Strips BOM, normalizes line endings, fixes common Windows-formatted HTML issues.
    #>
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Template file not found: $Path"
    }

    $encoding = Get-HtmlFileEncoding -Path $Path
    $content = [System.IO.File]::ReadAllText($Path, $encoding)

    # Strip BOM if present in the string
    if ($content.Length -gt 0 -and $content[0] -eq [char]0xFEFF) {
        $content = $content.Substring(1)
    }

    # Normalize line endings to \r\n (Outlook requirement)
    $content = $content -replace "`r?`n", "`r`n"

    # Fix common Windows-1252 artifacts that survived encoding conversion
    # Smart quotes → standard quotes (for email client compatibility)
    $content = $content -replace [char]0x201C, '&ldquo;'  # left double quote
    $content = $content -replace [char]0x201D, '&rdquo;'  # right double quote
    $content = $content -replace [char]0x2018, '&lsquo;'  # left single quote
    $content = $content -replace [char]0x2019, '&rsquo;'  # right single quote
    $content = $content -replace [char]0x2013, '&ndash;'  # en dash
    $content = $content -replace [char]0x2014, '&mdash;'  # em dash
    $content = $content -replace [char]0x2026, '&hellip;' # ellipsis
    $content = $content -replace [char]0x2022, '&bull;'   # bullet
    $content = $content -replace [char]0x20AC, '&euro;'   # euro sign
    $content = $content -replace [char]0x2122, '&trade;'  # trademark

    # Ensure charset meta tag is UTF-8
    if ($content -match 'charset\s*=\s*["\''']?\s*[\w-]+') {
        $content = $content -replace '(charset\s*=\s*["\''']?\s*)[\w-]+', '${1}utf-8'
    }

    Write-Log "  Template loaded: $Path ($($content.Length) chars, normalized to UTF-8)"
    return $content
}

#region ═══════════════════════════════════════════════════════════════════
# AUTHENTICATION — MSAL.PS or manual OAuth2
#endregion ════════════════════════════════════════════════════════════════

function Get-GraphToken {
    <#
    .SYNOPSIS
        Acquires an access token for Microsoft Graph.
        Supports: Client Credentials (app-only) or Device Code (delegated/interactive).
    #>

    $scopes = @(
        'https://graph.microsoft.com/MailboxConfigItem.ReadWrite'

        'https://graph.microsoft.com/User.Read.All'
        'https://graph.microsoft.com/MailboxSettings.ReadWrite'
    )

    if ($ClientSecret) {
        # App-only / client credentials flow
        Write-Log "Authenticating with client credentials (app-only) flow"
        $body = @{
            grant_type    = 'client_credentials'
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = 'https://graph.microsoft.com/.default'
        }
        $tokenResponse = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/${TenantId}/oauth2/v2.0/token" `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body $body

        $script:AccessToken = $tokenResponse.access_token
        $script:TokenExpires = (Get-Date).AddSeconds($tokenResponse.expires_in - 300)
        Write-Log "  Token acquired (app-only), expires in $($tokenResponse.expires_in)s" -Level SUCCESS
    }
    else {
        # Device Code flow (delegated / interactive)
        Write-Log "Authenticating with device code (delegated) flow"
        $deviceCodeBody = @{
            client_id = $ClientId
            scope     = ($scopes + 'offline_access openid profile email') -join ' '
        }
        $deviceCodeResponse = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/${TenantId}/oauth2/v2.0/devicecode" `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body $deviceCodeBody

        Write-Host ""
        Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║  AUTHENTICATION REQUIRED                                     ║" -ForegroundColor Cyan
        Write-Host "║  $($deviceCodeResponse.message.PadRight(60))║" -ForegroundColor Yellow
        Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""

        $pollBody = @{
            grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
            client_id   = $ClientId
            device_code = $deviceCodeResponse.device_code
        }

        $pollInterval = $deviceCodeResponse.interval
        $maxAttempts = [math]::Ceiling($deviceCodeResponse.expires_in / $pollInterval)
        $attempt = 0

        while ($attempt -lt $maxAttempts) {
            Start-Sleep -Seconds $pollInterval
            $attempt++
            try {
                $tokenResponse = Invoke-RestMethod -Method Post `
                    -Uri "https://login.microsoftonline.com/${TenantId}/oauth2/v2.0/token" `
                    -ContentType 'application/x-www-form-urlencoded' `
                    -Body $pollBody -ErrorAction Stop

                $script:AccessToken = $tokenResponse.access_token
                $script:TokenExpires = (Get-Date).AddSeconds($tokenResponse.expires_in - 300)
                Write-Log "  Token acquired (delegated), expires in $($tokenResponse.expires_in)s" -Level SUCCESS
                return
            }
            catch {
                $errorBody = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($errorBody.error -eq 'authorization_pending') {
                    Write-Host "." -NoNewline
                    continue
                }
                elseif ($errorBody.error -eq 'slow_down') {
                    $pollInterval += 5
                    continue
                }
                else {
                    throw "Authentication failed: $($errorBody.error_description)"
                }
            }
        }
        throw "Authentication timed out after $($deviceCodeResponse.expires_in) seconds."
    }
}

function Invoke-GraphRequest {
    <#
    .SYNOPSIS
        Wrapper for Graph API calls with automatic token refresh, retry, and throttling.
    #>
    param(
        [string]$Method = 'GET',
        [string]$Uri,
        [object]$Body,
        [string]$ContentType = 'application/json',
        [switch]$Beta,
        [int]$MaxRetries = 3
    )

    if (-not $script:AccessToken -or (Get-Date) -ge $script:TokenExpires) {
        Get-GraphToken
    }

    $baseUri = if ($Beta) { $script:GraphBeta } else { $script:GraphV1 }
    $fullUri = if ($Uri.StartsWith('http')) { $Uri } else { "${baseUri}${Uri}" }

    $headers = @{
        'Authorization' = "Bearer $($script:AccessToken)"
        'Content-Type'  = $ContentType
        'ConsistencyLevel' = 'eventual'
    }

    for ($retry = 0; $retry -lt $MaxRetries; $retry++) {
        try {
            $params = @{
                Method  = $Method
                Uri     = $fullUri
                Headers = $headers
            }
            if ($Body -and $Method -ne 'GET') {
                if ($Body -is [string]) {
                    $params['Body'] = $Body
                } else {
                    $params['Body'] = ($Body | ConvertTo-Json -Depth 10 -Compress)
                }
            }

            $response = Invoke-RestMethod @params -ErrorAction Stop
            return $response
        }
        catch {
            $statusCode = $_.Exception.Response.StatusCode.Value__
            if ($statusCode -eq 429) {
                # Throttled — respect Retry-After header
                $retryAfter = 30
                $retryHeader = $_.Exception.Response.Headers | Where-Object { $_.Key -eq 'Retry-After' }
                if ($retryHeader) { $retryAfter = [int]$retryHeader.Value[0] }
                Write-Log "  Throttled (429). Waiting ${retryAfter}s before retry $($retry+1)/$MaxRetries" -Level WARN
                Start-Sleep -Seconds $retryAfter
                continue
            }
            elseif ($statusCode -in @(500, 502, 503, 504) -and $retry -lt ($MaxRetries - 1)) {
                $wait = [math]::Pow(2, $retry) * 5
                Write-Log "  Server error ($statusCode). Retry $($retry+1)/$MaxRetries in ${wait}s" -Level WARN
                Start-Sleep -Seconds $wait
                continue
            }
            else {
                $errDetail = $_.ErrorDetails.Message
                Write-Log "  Graph API error: $statusCode — $errDetail" -Level ERROR
                throw
            }
        }
    }
}

#region ═══════════════════════════════════════════════════════════════════
# USER DATA — Graph API user properties for template variable replacement
#endregion ════════════════════════════════════════════════════════════════

function Get-UserProperties {
    param([string]$UPN)

    $select = @(
        'id', 'displayName', 'givenName', 'surname', 'mail', 'userPrincipalName',
        'jobTitle', 'department', 'companyName', 'officeLocation',
        'businessPhones', 'mobilePhone', 'faxNumber',
        'streetAddress', 'city', 'state', 'postalCode', 'country',
        'onPremisesExtensionAttributes', 'mailNickname', 'proxyAddresses'
    ) -join ','

    $user = Invoke-GraphRequest -Uri "/users/${UPN}?`$select=${select}" -Beta

    # Fetch manager
    try {
        $manager = Invoke-GraphRequest -Uri "/users/${UPN}/manager?`$select=displayName,mail,jobTitle" -Beta
        $user | Add-Member -NotePropertyName '_manager' -NotePropertyValue $manager -Force
    }
    catch {
        Write-Log "  No manager found for ${UPN}" -Level WARN
        $user | Add-Member -NotePropertyName '_manager' -NotePropertyValue $null -Force
    }

    return $user
}

function Expand-TemplateVariables {
    <#
    .SYNOPSIS
        Replaces {{Variable}} tokens in HTML with user properties from Graph.
        Supports nested properties, arrays, and custom extension attributes.
    #>
    param(
        [string]$Template,
        [psobject]$User
    )

    $result = $Template

    # Standard variable mappings
    $varMap = @{
        '{{DisplayName}}'    = $User.displayName
        '{{GivenName}}'      = $User.givenName
        '{{Surname}}'        = $User.surname
        '{{Mail}}'           = $User.mail
        '{{UPN}}'            = $User.userPrincipalName
        '{{JobTitle}}'       = $User.jobTitle
        '{{Department}}'     = $User.department
        '{{Company}}'        = $User.companyName
        '{{Office}}'         = $User.officeLocation
        '{{Phone}}'          = if ($User.businessPhones -and $User.businessPhones.Count -gt 0) { $User.businessPhones[0] } else { '' }
        '{{Phone2}}'         = if ($User.businessPhones -and $User.businessPhones.Count -gt 1) { $User.businessPhones[1] } else { '' }
        '{{Mobile}}'         = $User.mobilePhone
        '{{Fax}}'            = $User.faxNumber
        '{{Street}}'         = $User.streetAddress
        '{{City}}'           = $User.city
        '{{State}}'          = $User.state
        '{{PostalCode}}'     = $User.postalCode
        '{{Country}}'        = $User.country
        '{{MailNickname}}'   = $User.mailNickname
        '{{ManagerName}}'    = if ($User._manager) { $User._manager.displayName } else { '' }
        '{{ManagerMail}}'    = if ($User._manager) { $User._manager.mail } else { '' }
        '{{ManagerTitle}}'   = if ($User._manager) { $User._manager.jobTitle } else { '' }
    }

    # Extension attributes 1-15
    for ($i = 1; $i -le 15; $i++) {
        $attrName = "extensionAttribute$i"
        $value = ''
        if ($User.onPremisesExtensionAttributes -and $User.onPremisesExtensionAttributes.$attrName) {
            $value = $User.onPremisesExtensionAttributes.$attrName
        }
        $varMap["{{ExtAttr$i}}"] = $value
    }

    foreach ($key in $varMap.Keys) {
        $value = if ($null -eq $varMap[$key]) { '' } else { [System.Web.HttpUtility]::HtmlEncode($varMap[$key]) }
        $result = $result.Replace($key, $value)
    }

    # Remove any remaining unreplaced variables (clean output)
    # But keep track of what was missed
    $remaining = [regex]::Matches($result, '\{\{(\w+)\}\}')
    foreach ($match in $remaining) {
        Write-Log "  Unreplaced variable: $($match.Value) — removing from output" -Level WARN
        $result = $result.Replace($match.Value, '')
    }

    # Remove empty table rows (rows where ALL cells are empty after variable replacement)
    $result = $result -replace '<tr>\s*(<td[^>]*>\s*(<[^>]*>\s*)*\s*(</[^>]*>\s*)*\s*</td>\s*)+</tr>', ''

    # Remove empty lines caused by removed content
    $result = $result -replace '(\r?\n){3,}', "`r`n`r`n"

    return $result
}

#region ═══════════════════════════════════════════════════════════════════
# STRATEGY 1: ROAMING SIGNATURES via Graph Beta UserConfiguration API
#endregion ════════════════════════════════════════════════════════════════

function Set-RoamingSignature {
    <#
    .SYNOPSIS
        Deploys a signature as a roaming signature via the Graph Beta UserConfiguration API.
        Uses IPM.Configuration.OWA.UserOptions FAI in the Inbox for OWA settings,
        and MailboxConfigItem.ReadWrite for UserConfiguration access.

    .DESCRIPTION
        The roaming signature storage changed with New Outlook / Monarch.
        Old: Set-MailboxMessageConfiguration -SignatureHtml (requires PostponeRoamingSignaturesUntilLater)
        New: substrate.office.com outlookcloudsettings API (undocumented)
        Graph: userConfiguration API on /beta (MailboxConfigItem.ReadWrite) — since Jan 2026

        This function writes the OWA.UserOptions UserConfiguration object which controls
        the classic OWA signature. For roaming signatures (New Outlook / Monarch), the
        substrate API is used internally by Outlook itself. The Graph approach here sets
        the classic OWA path, which is still the most reliable programmatic approach.

        For full roaming signature support equivalent to Set-OutlookSignatures Benefactor Circle,
        the substrate.office.com API would need to be reverse-engineered (not recommended for production).
    #>
    param(
        [string]$UserUPN,
        [string]$SignatureHtml,
        [string]$SignatureText,
        [string]$SignatureName,
        [bool]$SetAsDefault = $true,
        [bool]$SetForReply = $true
    )

    Write-Log "  [Roaming] Setting signature for ${UserUPN} via Graph Beta UserConfiguration API"

    # Step 1: Get the Inbox folder ID
    # inbox is a well-known folder name — no Mail.ReadWrite needed

    # Step 2: Try to read existing OWA.UserOptions UserConfiguration
    $configName = 'OWA.UserOptions'
    $existingConfig = $null

    try {
        $existingConfig = Invoke-GraphRequest `
            -Uri "/users/${UserUPN}/mailFolders/inbox/userConfigurations/${configName}" `
            -Beta
        Write-Log "    Existing UserConfiguration found: $configName"
    }
    catch {
        Write-Log "    No existing UserConfiguration '$configName' — will use PATCH on mailboxSettings" -Level WARN
    }

    if ($existingConfig) {
        # Decode existing structuredData (RoamingDictionary)
        $structuredData = @{}
        if ($existingConfig.structuredData) {
            try {
                $rawBytes = [Convert]::FromBase64String($existingConfig.structuredData)
                $decoded = [System.Text.Encoding]::UTF8.GetString($rawBytes)
                # The RoamingDictionary is XML-based — parse it
                # For simplicity, we'll work with the known keys
                Write-Log "    Decoded structuredData ($($rawBytes.Length) bytes)"
            }
            catch {
                Write-Log "    Could not decode structuredData: $_" -Level WARN
            }
        }

        # Build XML dictionary for OWA UserOptions
        # Known keys for signatures: signaturehtml, signaturetext, autoaddsignature, autoaddsignatureonreply
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
      <DictionaryValue><Type>Boolean</Type><Value>$(if($SetAsDefault){'true'}else{'false'})</Value></DictionaryValue>
    </DictionaryEntry>
    <DictionaryEntry>
      <DictionaryKey><Type>String</Type><Value>autoaddsignatureonreply</Value></DictionaryKey>
      <DictionaryValue><Type>Boolean</Type><Value>$(if($SetForReply){'true'}else{'false'})</Value></DictionaryValue>
    </DictionaryEntry>
  </Dictionary>
</UserConfiguration>
"@

        $base64Dict = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($dictXml))

        $updateBody = @{
            '@odata.type' = '#microsoft.graph.userConfiguration'
            structuredData = $base64Dict
        }

        if (-not $DryRun) {
            try {
                Invoke-GraphRequest `
                    -Method PATCH `
                    -Uri "/users/${UserUPN}/mailFolders/inbox/userConfigurations/${configName}" `
                    -Body $updateBody `
                    -Beta
                Write-Log "    UserConfiguration '$configName' updated successfully" -Level SUCCESS
            }
            catch {
                Write-Log "    PATCH failed, trying PUT: $_" -Level WARN
                # Some configurations require PUT instead of PATCH
                Invoke-GraphRequest `
                    -Method PUT `
                    -Uri "/users/${UserUPN}/mailFolders/inbox/userConfigurations/${configName}" `
                    -Body $updateBody `
                    -Beta
                Write-Log "    UserConfiguration '$configName' replaced via PUT" -Level SUCCESS
            }
        }
        else {
            Write-Log "    [DRY RUN] Would update UserConfiguration '$configName'" -Level WARN
        }
    }
    else {
        # Fallback: Use EXO REST InvokeCommand to call Set-MailboxMessageConfiguration
        # This works when the UserConfiguration doesn't exist or can't be accessed
        Write-Log "    Falling back to EXO InvokeCommand for Set-MailboxMessageConfiguration"
        Set-OWASignatureViaEXORest -UserUPN $UserUPN -SignatureHtml $SignatureHtml `
            -SignatureText $SignatureText -SetAsDefault:$SetAsDefault -SetForReply:$SetForReply
    }
}

#region ═══════════════════════════════════════════════════════════════════
# STRATEGY 2: OWA CLASSIC via EXO REST InvokeCommand
#endregion ════════════════════════════════════════════════════════════════

function Get-EXOToken {
    <#
    .SYNOPSIS
        Gets a token scoped to Office 365 Exchange Online for InvokeCommand endpoint.
    #>
    $body = @{
        grant_type    = 'client_credentials'
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'https://outlook.office365.com/.default'
    }

    if (-not $ClientSecret) {
        Write-Log "  EXO REST requires client_secret for app-only InvokeCommand. Use -ClientSecret or Strategy 'Roaming'." -Level ERROR
        return $null
    }

    $tokenResponse = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/${TenantId}/oauth2/v2.0/token" `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body $body

    return $tokenResponse.access_token
}

function Set-OWASignatureViaEXORest {
    <#
    .SYNOPSIS
        Sets OWA signature via EXO REST InvokeCommand endpoint.
        This is the same mechanism ExchangeOnlineManagement v3+ uses internally.
        IMPORTANT: Only works if PostponeRoamingSignaturesUntilLater is still $true in your org,
        OR for tenants where roaming signatures haven't fully taken over.
    #>
    param(
        [string]$UserUPN,
        [string]$SignatureHtml,
        [string]$SignatureText,
        [bool]$SetAsDefault = $true,
        [bool]$SetForReply = $true
    )

    $exoToken = Get-EXOToken
    if (-not $exoToken) {
        Write-Log "    Cannot obtain EXO token — skipping OWA fallback" -Level ERROR
        return
    }

    $payload = @{
        CmdletInput = @{
            CmdletName = 'Set-MailboxMessageConfiguration'
            Parameters = @{
                Identity                       = $UserUPN
                SignatureHtml                   = $SignatureHtml
                SignatureText                   = $SignatureText
                AutoAddSignature               = $SetAsDefault
                AutoAddSignatureOnReply         = $SetForReply
                ErrorAction                    = 'Stop'
            }
        }
    }

    $headers = @{
        'Authorization'        = "Bearer $exoToken"
        'Content-Type'         = 'application/json'
        'X-CmdletName'         = 'Set-MailboxMessageConfiguration'
        'X-ResponseFormat'     = 'json'
        'X-ClientApplication'  = 'SignatureManager/1.0'
        'X-AnchorMailbox'      = "UPN:${UserUPN}"
    }

    if (-not $DryRun) {
        try {
            $response = Invoke-RestMethod -Method Post `
                -Uri "https://outlook.office365.com/adminapi/beta/${TenantId}/InvokeCommand" `
                -Headers $headers `
                -Body ($payload | ConvertTo-Json -Depth 10) `
                -ErrorAction Stop

            Write-Log "    OWA signature set via EXO REST InvokeCommand" -Level SUCCESS
        }
        catch {
            $errMsg = $_.ErrorDetails.Message
            if ($errMsg -match 'roaming') {
                Write-Log "    EXO InvokeCommand rejected — roaming signatures active. Use Strategy 'Roaming' or set PostponeRoamingSignaturesUntilLater." -Level ERROR
            }
            else {
                Write-Log "    EXO InvokeCommand failed: $errMsg" -Level ERROR
            }
        }
    }
    else {
        Write-Log "    [DRY RUN] Would call Set-MailboxMessageConfiguration for ${UserUPN}" -Level WARN
    }
}

#region ═══════════════════════════════════════════════════════════════════
# STRATEGY 3: TRANSPORT RULE — server-side injection
#endregion ════════════════════════════════════════════════════════════════

function Set-TransportRuleSignature {
    <#
    .SYNOPSIS
        Creates/updates an Exchange Online transport rule that appends signatures server-side.
        Works for ALL clients (Outlook, OWA, mobile, third-party).
        Limitation: No per-user customization unless combined with DLP/user conditions.

    .DESCRIPTION
        Uses the EXO REST InvokeCommand for New-TransportRule / Set-TransportRule.
        The %%DisplayName%%, %%Email%% etc. tokens are Exchange transport rule variables.
    #>
    param(
        [string]$SignatureHtml,
        [string]$RuleName = 'Corporate Email Signature'
    )

    Write-Log "[TransportRule] Creating/updating transport rule: $RuleName"

    # Exchange transport rules support these built-in variables:
    # %%displayName%%, %%email%%, %%title%%, %%department%%, %%company%%,
    # %%phoneNumber%%, %%mobileNumber%%, %%faxNumber%%,
    # %%street%%, %%city%%, %%state%%, %%zipCode%%, %%country%%,
    # %%officeLocation%%, %%customAttribute1-15%%
    # Convert our template variables to Exchange transport rule format
    $trHtml = $SignatureHtml
    $exchangeVarMap = @{
        '{{DisplayName}}'  = '%%displayName%%'
        '{{GivenName}}'    = '%%firstName%%'
        '{{Surname}}'      = '%%lastName%%'
        '{{Mail}}'         = '%%email%%'
        '{{JobTitle}}'     = '%%title%%'
        '{{Department}}'   = '%%department%%'
        '{{Company}}'      = '%%company%%'
        '{{Phone}}'        = '%%phoneNumber%%'
        '{{Mobile}}'       = '%%mobileNumber%%'
        '{{Fax}}'          = '%%faxNumber%%'
        '{{Street}}'       = '%%street%%'
        '{{City}}'         = '%%city%%'
        '{{State}}'        = '%%state%%'
        '{{PostalCode}}'   = '%%zipCode%%'
        '{{Country}}'      = '%%country%%'
        '{{Office}}'       = '%%officeLocation%%'
    }
    for ($i = 1; $i -le 15; $i++) {
        $exchangeVarMap["{{ExtAttr$i}}"] = "%%customAttribute${i}%%"
    }

    foreach ($key in $exchangeVarMap.Keys) {
        $trHtml = $trHtml.Replace($key, $exchangeVarMap[$key])
    }

    # Wrap in disclaimer format
    $disclaimerHtml = "<hr style='border:none;border-top:1px solid #e0e0e0;margin:16px 0;'/>" + $trHtml

    $exoToken = Get-EXOToken
    if (-not $exoToken) {
        Write-Log "  Cannot obtain EXO token for transport rule" -Level ERROR
        return
    }

    # Check if rule exists
    $checkPayload = @{
        CmdletInput = @{
            CmdletName = 'Get-TransportRule'
            Parameters = @{
                Identity    = $RuleName
                ErrorAction = 'SilentlyContinue'
            }
        }
    }
    $headers = @{
        'Authorization'        = "Bearer $exoToken"
        'Content-Type'         = 'application/json'
        'X-CmdletName'         = 'Get-TransportRule'
        'X-ResponseFormat'     = 'json'
        'X-ClientApplication'  = 'SignatureManager/1.0'
    }

    $ruleExists = $false
    try {
        $existing = Invoke-RestMethod -Method Post `
            -Uri "https://outlook.office365.com/adminapi/beta/${TenantId}/InvokeCommand" `
            -Headers $headers -Body ($checkPayload | ConvertTo-Json -Depth 10) -ErrorAction Stop
        if ($existing) { $ruleExists = $true }
    }
    catch { }

    $cmdlet = if ($ruleExists) { 'Set-TransportRule' } else { 'New-TransportRule' }
    $ruleParams = @{
        Name                       = $RuleName
        ApplyHtmlDisclaimerText    = $disclaimerHtml
        ApplyHtmlDisclaimerLocation = 'Append'
        ApplyHtmlDisclaimerFallbackAction = 'Wrap'
        FromScope                  = 'InOrganization'
        SentToScope                = 'NotInOrganization'
        ErrorAction                = 'Stop'
    }
    if ($ruleExists) {
        $ruleParams.Remove('Name')
        $ruleParams['Identity'] = $RuleName
    }

    $createPayload = @{
        CmdletInput = @{
            CmdletName = $cmdlet
            Parameters = $ruleParams
        }
    }
    $headers['X-CmdletName'] = $cmdlet

    if (-not $DryRun) {
        try {
            Invoke-RestMethod -Method Post `
                -Uri "https://outlook.office365.com/adminapi/beta/${TenantId}/InvokeCommand" `
                -Headers $headers -Body ($createPayload | ConvertTo-Json -Depth 10) -ErrorAction Stop
            Write-Log "  Transport rule '$RuleName' $( if($ruleExists){'updated'}else{'created'} ) successfully" -Level SUCCESS
        }
        catch {
            Write-Log "  Transport rule $cmdlet failed: $($_.ErrorDetails.Message)" -Level ERROR
        }
    }
    else {
        Write-Log "  [DRY RUN] Would $cmdlet transport rule '$RuleName'" -Level WARN
    }
}

#region ═══════════════════════════════════════════════════════════════════
# HTML PROCESSING HELPERS
#endregion ════════════════════════════════════════════════════════════════

function ConvertTo-PlainText {
    <#
    .SYNOPSIS
        Converts HTML to plain text, preserving basic structure.
    #>
    param([string]$Html)

    $text = $Html
    $text = $text -replace '<br\s*/?>', "`r`n"
    $text = $text -replace '</p>', "`r`n"
    $text = $text -replace '</div>', "`r`n"
    $text = $text -replace '</tr>', "`r`n"
    $text = $text -replace '</li>', "`r`n"
    $text = $text -replace '<hr[^>]*>', ('-' * 40 + "`r`n")
    $text = $text -replace '<[^>]+>', ''
    $text = [System.Web.HttpUtility]::HtmlDecode($text)
    $text = $text -replace '(\r?\n){3,}', "`r`n`r`n"
    return $text.Trim()
}

function Assert-HtmlSafe {
    <#
    .SYNOPSIS
        Strips potentially dangerous content from HTML signatures.
        Removes script tags, event handlers, external links to non-image resources.
    #>
    param([string]$Html)

    $safe = $Html
    $safe = $safe -replace '<script[^>]*>[\s\S]*?</script>', ''
    $safe = $safe -replace '\son\w+\s*=\s*["\'''][^""\'']*["\''']', ''
    $safe = $safe -replace '<iframe[^>]*>[\s\S]*?</iframe>', ''
    $safe = $safe -replace '<object[^>]*>[\s\S]*?</object>', ''
    $safe = $safe -replace '<embed[^>]*>', ''
    $safe = $safe -replace '<form[^>]*>[\s\S]*?</form>', ''
    $safe = $safe -replace '<input[^>]*>', ''
    $safe = $safe -replace 'javascript:', ''
    $safe = $safe -replace 'vbscript:', ''
    return $safe
}

#region ═══════════════════════════════════════════════════════════════════
# MAIN ORCHESTRATOR
#endregion ════════════════════════════════════════════════════════════════

function Main {
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║  Manage-OutlookSignatures — Enterprise Signature Deployment      ║" -ForegroundColor Cyan
    Write-Host "║  Strategy: $($Strategy.PadRight(52))║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""

    Write-Log "Starting signature deployment — Strategy: $Strategy"
    Write-Log "Template: $TemplatePath"
    Write-Log "Target: $(if($UserUPN -eq '*'){'All licensed users'}else{$UserUPN})"
    if ($DryRun) { Write-Log "*** DRY RUN MODE — no changes will be made ***" -Level WARN }

    # Load and validate template
    Write-Log "Loading HTML template..."
    $templateHtml = Read-HtmlTemplate -Path $TemplatePath
    $templateHtml = Assert-HtmlSafe -Html $templateHtml
    Write-Log "Template loaded and sanitized ($($templateHtml.Length) chars)"

    # Authenticate
    Get-GraphToken

    # Resolve target users
    $targetUsers = @()
    if ($UserUPN -eq '*') {
        Write-Log "Fetching all licensed users..."
        $response = Invoke-GraphRequest -Uri "/users?`$filter=assignedLicenses/`$count ne 0&`$count=true&`$select=userPrincipalName,displayName,mail&`$top=999" -Beta
        $targetUsers = $response.value | Where-Object { $_.mail }
        Write-Log "Found $($targetUsers.Count) licensed users with mailboxes"

        # Handle paging
        while ($response.'@odata.nextLink') {
            $response = Invoke-GraphRequest -Uri $response.'@odata.nextLink'
            $targetUsers += $response.value | Where-Object { $_.mail }
        }
    }
    else {
        $targetUsers = @(@{ userPrincipalName = $UserUPN; displayName = $UserUPN })
    }

    # Process each user
    $successCount = 0
    $errorCount = 0
    $total = $targetUsers.Count

    for ($i = 0; $i -lt $total; $i++) {
        $targetUser = $targetUsers[$i]
        $upn = $targetUser.userPrincipalName
        $pct = [math]::Round((($i + 1) / $total) * 100)
        Write-Log "[$($i+1)/${total}] (${pct}%) Processing: $upn"

        try {
            # Fetch full user properties for variable replacement
            $userProps = Get-UserProperties -UPN $upn

            # Expand template variables
            $personalizedHtml = Expand-TemplateVariables -Template $templateHtml -User $userProps
            $personalizedText = ConvertTo-PlainText -Html $personalizedHtml

            Write-Log "  Signature generated: $($personalizedHtml.Length) chars HTML, $($personalizedText.Length) chars text"

            # Deploy based on strategy
            switch ($Strategy) {
                'Roaming' {
                    Set-RoamingSignature -UserUPN $upn -SignatureHtml $personalizedHtml `
                        -SignatureText $personalizedText -SignatureName $SignatureName `
                        -SetAsDefault:$SetAsDefault -SetForReply:$SetForReply
                }
                'OWA' {
                    Set-OWASignatureViaEXORest -UserUPN $upn -SignatureHtml $personalizedHtml `
                        -SignatureText $personalizedText -SetAsDefault:$SetAsDefault -SetForReply:$SetForReply
                }
                'TransportRule' {
                    # Transport rules are org-wide, so we only need to run once with the template
                    Set-TransportRuleSignature -SignatureHtml $templateHtml
                    Write-Log "Transport rule created — applies to all users" -Level SUCCESS
                    return  # Done, no need to iterate users
                }
                'All' {
                    Set-RoamingSignature -UserUPN $upn -SignatureHtml $personalizedHtml `
                        -SignatureText $personalizedText -SignatureName $SignatureName `
                        -SetAsDefault:$SetAsDefault -SetForReply:$SetForReply
                }
            }

            $successCount++
        }
        catch {
            $errorCount++
            Write-Log "  FAILED for ${upn}: $_" -Level ERROR
        }
    }

    # Summary
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║  DEPLOYMENT COMPLETE                                             ║" -ForegroundColor Cyan
    Write-Host "║  Success: $($successCount.ToString().PadRight(5)) Failed: $($errorCount.ToString().PadRight(38))║" -ForegroundColor $(if($errorCount -gt 0){'Yellow'}else{'Green'})
    Write-Host "║  Log: $($LogPath.PadRight(57))║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
}

# ═══════════════════════════════════════════════════════════════════════════
# RUN
# ═══════════════════════════════════════════════════════════════════════════
Add-Type -AssemblyName System.Web
Main
