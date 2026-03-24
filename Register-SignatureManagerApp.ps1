<#
.SYNOPSIS
    Register-SignatureManagerApp.ps1 — Creates the Entra ID App Registration
    with all required permissions for Manage-OutlookSignatures.ps1

.DESCRIPTION
    Creates a single-tenant app with:
      - MailboxConfigItem.ReadWrite (Graph Beta — roaming signatures)
      - Mail.ReadWrite (Graph — mailbox access)
      - User.Read.All (Graph — user properties)
      - MailboxSettings.ReadWrite (Graph — OOF and mailbox settings)
      - Exchange.ManageAsApp (Office 365 Exchange Online — EXO REST)

    Supports delegated (device code / interactive) and app-only (client secret) flows.

.PARAMETER AppName
    Display name for the app registration.

.PARAMETER CreateClientSecret
    Whether to create a client secret for app-only / daemon scenarios.

.EXAMPLE
    # Interactive — will prompt for Global Admin / App Admin credentials
    .\Register-SignatureManagerApp.ps1 -AppName "Signature Manager"
#>

[CmdletBinding()]
param(
    [string]$AppName = 'Signature Manager — Contoso',
    [switch]$CreateClientSecret,
    [int]$SecretExpiryDays = 365
)

$ErrorActionPreference = 'Stop'

# ═══════════════════════════════════════════════════════════════════
# Check prerequisites
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n  Checking prerequisites..." -ForegroundColor Cyan

$mgModule = Get-Module -ListAvailable -Name Microsoft.Graph.Applications
if (-not $mgModule) {
    Write-Host "  Installing Microsoft.Graph modules..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force
    Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser -Force
}

# ═══════════════════════════════════════════════════════════════════
# Connect to Graph with admin permissions
# ═══════════════════════════════════════════════════════════════════
Write-Host "  Connecting to Microsoft Graph (admin consent required)..." -ForegroundColor Cyan
Connect-MgGraph -Scopes @(
    'Application.ReadWrite.All'
    'AppRoleAssignment.ReadWrite.All'
    'DelegatedPermissionGrant.ReadWrite.All'
) -NoWelcome

$context = Get-MgContext
$tenantId = $context.TenantId
Write-Host "  Connected as: $($context.Account) in tenant: $tenantId" -ForegroundColor Green

# ═══════════════════════════════════════════════════════════════════
# Define required permissions
# ═══════════════════════════════════════════════════════════════════

# Microsoft Graph App ID (well-known)
$graphAppId = '00000003-0000-0000-c000-000000000000'
# Office 365 Exchange Online App ID (well-known)
$exoAppId   = '00000002-0000-0ff1-ce00-000000000000'

# Get service principals
$graphSP = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"
$exoSP   = Get-MgServicePrincipal -Filter "appId eq '$exoAppId'"

# Graph delegated permissions
$graphDelegatedPermissions = @(
    'MailboxConfigItem.ReadWrite'   # Roaming signatures via UserConfiguration API
    'Mail.ReadWrite'                 # Mailbox access
    'User.Read.All'                  # User properties for variable replacement
    'MailboxSettings.ReadWrite'      # OOF messages and mailbox settings
    'offline_access'                 # Refresh token
    'openid'                         # Sign-in
    'profile'                        # Basic profile
    'email'                          # Email address
    'GroupMember.Read.All'           # Group membership for assignment rules
)

# Graph application permissions (for app-only/daemon flow)
$graphAppPermissions = @(
    'MailboxConfigItem.ReadWrite'
    'Mail.ReadWrite'
    'User.Read.All'
    'MailboxSettings.ReadWrite'
)

# EXO application permission (for InvokeCommand / transport rules)
$exoAppPermissions = @(
    'Exchange.ManageAsApp'
)

# Resolve permission IDs
function Resolve-PermissionIds {
    param($ServicePrincipal, $Permissions, $Type)
    $resolved = @()
    foreach ($permName in $Permissions) {
        if ($Type -eq 'Delegated') {
            $perm = $ServicePrincipal.Oauth2PermissionScopes | Where-Object { $_.Value -eq $permName }
        }
        else {
            $perm = $ServicePrincipal.AppRoles | Where-Object { $_.Value -eq $permName }
        }
        if ($perm) {
            $resolved += @{
                Id   = $perm.Id
                Type = if ($Type -eq 'Delegated') { 'Scope' } else { 'Role' }
                Name = $permName
            }
            Write-Host "    [OK] $permName ($Type)" -ForegroundColor DarkGray
        }
        else {
            Write-Host "    [SKIP] $permName — not found (may require beta/preview)" -ForegroundColor Yellow
        }
    }
    return $resolved
}

Write-Host "`n  Resolving Graph permissions..." -ForegroundColor Cyan
$graphDelegatedResolved = Resolve-PermissionIds -ServicePrincipal $graphSP -Permissions $graphDelegatedPermissions -Type 'Delegated'
$graphAppResolved       = Resolve-PermissionIds -ServicePrincipal $graphSP -Permissions $graphAppPermissions -Type 'Application'

Write-Host "`n  Resolving Exchange Online permissions..." -ForegroundColor Cyan
$exoAppResolved = Resolve-PermissionIds -ServicePrincipal $exoSP -Permissions $exoAppPermissions -Type 'Application'

# ═══════════════════════════════════════════════════════════════════
# Create App Registration
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n  Creating app registration: '$AppName'..." -ForegroundColor Cyan

# Build required resource access
$resourceAccess = @()

# Graph permissions
$graphAccess = @{
    ResourceAppId  = $graphAppId
    ResourceAccess = @()
}
foreach ($p in $graphDelegatedResolved) {
    $graphAccess.ResourceAccess += @{ Id = $p.Id; Type = $p.Type }
}
foreach ($p in $graphAppResolved) {
    $graphAccess.ResourceAccess += @{ Id = $p.Id; Type = $p.Type }
}
$resourceAccess += $graphAccess

# EXO permissions
if ($exoAppResolved.Count -gt 0) {
    $exoAccess = @{
        ResourceAppId  = $exoAppId
        ResourceAccess = @()
    }
    foreach ($p in $exoAppResolved) {
        $exoAccess.ResourceAccess += @{ Id = $p.Id; Type = $p.Type }
    }
    $resourceAccess += $exoAccess
}

$appParams = @{
    DisplayName            = $AppName
    SignInAudience         = 'AzureADMyOrg'  # Single tenant
    RequiredResourceAccess = $resourceAccess
    PublicClient           = @{
        RedirectUris = @(
            'http://localhost'
            'https://login.microsoftonline.com/common/oauth2/nativeclient'
        )
    }
    Web                    = @{
        RedirectUris = @('http://localhost')
    }
    IsFallbackPublicClient = $true  # Enable public client flows (device code, etc.)
}

$app = New-MgApplication @appParams
$appId = $app.AppId
$appObjectId = $app.Id
Write-Host "  App created: $AppName" -ForegroundColor Green
Write-Host "    Application (client) ID : $appId" -ForegroundColor White
Write-Host "    Object ID               : $appObjectId" -ForegroundColor White

# ═══════════════════════════════════════════════════════════════════
# Create Service Principal
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n  Creating service principal..." -ForegroundColor Cyan
$sp = New-MgServicePrincipal -AppId $appId
Write-Host "  Service principal created: $($sp.Id)" -ForegroundColor Green

# ═══════════════════════════════════════════════════════════════════
# Create Client Secret (optional)
# ═══════════════════════════════════════════════════════════════════
$clientSecret = $null
if ($CreateClientSecret) {
    Write-Host "`n  Creating client secret (expires in $SecretExpiryDays days)..." -ForegroundColor Cyan
    $secretParams = @{
        PasswordCredential = @{
            DisplayName = "Signature Manager Secret"
            EndDateTime = (Get-Date).AddDays($SecretExpiryDays)
        }
    }
    $secret = Add-MgApplicationPassword -ApplicationId $appObjectId @secretParams
    $clientSecret = $secret.SecretText
    Write-Host "  Client secret created" -ForegroundColor Green
    Write-Host "  *** SAVE THIS NOW — it cannot be retrieved later ***" -ForegroundColor Red
    Write-Host "    Secret: $clientSecret" -ForegroundColor Yellow
}

# ═══════════════════════════════════════════════════════════════════
# Grant Admin Consent
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n  Granting admin consent for application permissions..." -ForegroundColor Cyan

foreach ($p in $graphAppResolved) {
    try {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id `
            -ResourceId $graphSP.Id -AppRoleId $p.Id -ErrorAction Stop | Out-Null
        Write-Host "    [OK] Graph: $($p.Name)" -ForegroundColor Green
    }
    catch {
        Write-Host "    [WARN] Graph: $($p.Name) — $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

foreach ($p in $exoAppResolved) {
    try {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id `
            -ResourceId $exoSP.Id -AppRoleId $p.Id -ErrorAction Stop | Out-Null
        Write-Host "    [OK] EXO: $($p.Name)" -ForegroundColor Green
    }
    catch {
        Write-Host "    [WARN] EXO: $($p.Name) — $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# For delegated permissions, open the consent URL
$consentUrl = "https://login.microsoftonline.com/${tenantId}/adminconsent?client_id=${appId}"
Write-Host "`n  For delegated permissions, grant admin consent via:" -ForegroundColor Cyan
Write-Host "    $consentUrl" -ForegroundColor White

# ═══════════════════════════════════════════════════════════════════
# Output summary
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n╔══════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  APP REGISTRATION COMPLETE                                       ║" -ForegroundColor Green
Write-Host "╠══════════════════════════════════════════════════════════════════╣" -ForegroundColor Green
Write-Host "║  Tenant ID    : $($tenantId.PadRight(47))║" -ForegroundColor White
Write-Host "║  Client ID    : $($appId.PadRight(47))║" -ForegroundColor White
if ($clientSecret) {
    Write-Host "║  Secret       : $($clientSecret.Substring(0,[Math]::Min(20,$clientSecret.Length)).PadRight(47))║" -ForegroundColor Yellow
}
Write-Host "╠══════════════════════════════════════════════════════════════════╣" -ForegroundColor Green
Write-Host "║  Usage:                                                          ║" -ForegroundColor Cyan
Write-Host "║  # Delegated (interactive)                                       ║" -ForegroundColor DarkGray
Write-Host "║  .\Manage-OutlookSignatures.ps1 -TenantId `"$tenantId`"          ║" -ForegroundColor DarkGray
Write-Host "║      -ClientId `"$appId`"                                         ║" -ForegroundColor DarkGray
Write-Host "║      -TemplatePath `".\templates\corporate.htm`"                  ║" -ForegroundColor DarkGray
if ($clientSecret) {
    Write-Host "║  # App-only (daemon)                                             ║" -ForegroundColor DarkGray
    Write-Host "║  .\Manage-OutlookSignatures.ps1 -TenantId `"...`"                ║" -ForegroundColor DarkGray
    Write-Host "║      -ClientId `"...`" -ClientSecret `"...`"                       ║" -ForegroundColor DarkGray
}
Write-Host "╚══════════════════════════════════════════════════════════════════╝" -ForegroundColor Green

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
