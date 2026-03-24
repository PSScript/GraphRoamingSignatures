# Manage-OutlookSignatures

Enterprise email signature management for Microsoft 365 / Exchange Online via Microsoft Graph API.

## Architecture & Strategy

Microsoft has **no official Graph API** for roaming signatures as of March 2026. This solution uses three complementary strategies:

### Strategy 1: `Roaming` (recommended)
Uses the **Graph Beta UserConfiguration API** (`MailboxConfigItem.ReadWrite`) released January 2026.
Writes the `OWA.UserOptions` FAI (Folder Associated Item) in the Inbox to set signatures.

**Pro**: Works with New Outlook, OWA, and can coexist with roaming signatures.
**Con**: Beta API, may change. Does not directly write substrate cloud settings (New Outlook reads from there).

### Strategy 2: `OWA` (legacy fallback)
Uses the **EXO REST InvokeCommand** endpoint to call `Set-MailboxMessageConfiguration`.

**Pro**: Well-known, stable.
**Con**: Only works if `PostponeRoamingSignaturesUntilLater = $true` (deprecated path).

### Strategy 3: `TransportRule` (server-side)
Creates Exchange Online transport rules that append signatures to outbound emails.

**Pro**: Works for ALL clients (Outlook, OWA, mobile, third-party). No client-side deployment.
**Con**: No per-user personalization beyond Exchange built-in variables. Appended, not composed in client.

## Prerequisites

### Entra ID App Registration
Run `Register-SignatureManagerApp.ps1` to create the app with all required permissions:

```powershell
.\Register-SignatureManagerApp.ps1 -AppName "Signature Manager" -CreateClientSecret
```

Required Graph permissions (delegated + application):

| Permission | Type | Purpose |
|---|---|---|
| `MailboxConfigItem.ReadWrite` | Delegated + App | Roaming signatures via UserConfiguration API |
| `Mail.ReadWrite` | Delegated + App | Mailbox access |
| `User.Read.All` | Delegated + App | User properties for template variables |
| `MailboxSettings.ReadWrite` | Delegated + App | OOF and mailbox settings |
| `Exchange.ManageAsApp` | Application | EXO InvokeCommand (transport rules, OWA fallback) |

### PowerShell
- PowerShell 7.x+ recommended (5.1 works for most features)
- No external modules required — uses raw REST API calls
- `Microsoft.Graph.Applications` only needed for `Register-SignatureManagerApp.ps1`

## Usage

### Single user (interactive / delegated)
```powershell
.\Manage-OutlookSignatures.ps1 `
    -TenantId "your-tenant-id" `
    -ClientId "your-client-id" `
    -TemplatePath ".\templates\corporate.htm" `
    -UserUPN "user@contoso.com" `
    -Strategy Roaming `
    -SetAsDefault -SetForReply
```

### All users (app-only / daemon)
```powershell
.\Manage-OutlookSignatures.ps1 `
    -TenantId "your-tenant-id" `
    -ClientId "your-client-id" `
    -ClientSecret "your-secret" `
    -TemplatePath ".\templates\corporate.htm" `
    -Strategy Roaming `
    -SetAsDefault -SetForReply
```

### Transport rule (server-side for all users)
```powershell
.\Manage-OutlookSignatures.ps1 `
    -TenantId "your-tenant-id" `
    -ClientId "your-client-id" `
    -ClientSecret "your-secret" `
    -TemplatePath ".\templates\corporate.htm" `
    -Strategy TransportRule
```

### Dry run (preview without changes)
```powershell
.\Manage-OutlookSignatures.ps1 `
    -TenantId "..." -ClientId "..." `
    -TemplatePath ".\templates\corporate.htm" `
    -UserUPN "user@contoso.com" `
    -Strategy Roaming -DryRun
```

## Template Variables

Templates use `{{Variable}}` syntax. Variables are replaced with user data from Microsoft Graph.

| Variable | Graph Property | Description |
|---|---|---|
| `{{DisplayName}}` | `displayName` | Full name |
| `{{GivenName}}` | `givenName` | First name |
| `{{Surname}}` | `surname` | Last name |
| `{{JobTitle}}` | `jobTitle` | Job title |
| `{{Department}}` | `department` | Department |
| `{{Mail}}` | `mail` | Primary email |
| `{{Phone}}` | `businessPhones[0]` | Business phone |
| `{{Phone2}}` | `businessPhones[1]` | Second business phone |
| `{{Mobile}}` | `mobilePhone` | Mobile phone |
| `{{Fax}}` | `faxNumber` | Fax number |
| `{{Company}}` | `companyName` | Company name |
| `{{Office}}` | `officeLocation` | Office location |
| `{{Street}}` | `streetAddress` | Street address |
| `{{City}}` | `city` | City |
| `{{State}}` | `state` | State / province |
| `{{PostalCode}}` | `postalCode` | Postal code |
| `{{Country}}` | `country` | Country |
| `{{ManagerName}}` | `manager.displayName` | Manager's name |
| `{{ManagerMail}}` | `manager.mail` | Manager's email |
| `{{ManagerTitle}}` | `manager.jobTitle` | Manager's title |
| `{{ExtAttr1}}` – `{{ExtAttr15}}` | `onPremisesExtensionAttributes` | Extension attributes 1-15 |

### Transport Rule Variables
When using `Strategy TransportRule`, the `{{Variable}}` tokens are automatically converted to Exchange transport rule variables (`%%variable%%`).

## Encoding Handling

The script automatically detects and handles:

| Encoding | Detection | Notes |
|---|---|---|
| **UTF-8** | BOM or default | Recommended for new templates |
| **UTF-8 BOM** | Byte sequence `EF BB BF` | BOM is stripped during processing |
| **Windows-1252** | Charset meta tag or byte heuristic (0x80-0x9F range) | Common in templates created with MS Word / Outlook |
| **ISO-8859-1** | Charset meta tag or UTF-8 validation failure | Western European fallback |
| **UTF-16 LE/BE** | BOM detection | Rare in HTML, fully supported |

All templates are normalized to UTF-8 during processing. Smart quotes, em dashes, and other Windows-1252 specific characters are converted to HTML entities for maximum email client compatibility.

## File Structure

```
signature-manager/
├── Manage-OutlookSignatures.ps1       # Main deployment script
├── Register-SignatureManagerApp.ps1    # Entra ID app registration
├── templates/
│   └── corporate.htm                  # Sample corporate signature template
└── README.md                          # This file
```

## Relationship to Set-OutlookSignatures

This tool complements [Set-OutlookSignatures](https://github.com/Set-OutlookSignatures/Set-OutlookSignatures) (the gold standard for PowerShell-based signature management). Key differences:

| Feature | This Tool | Set-OutlookSignatures |
|---|---|---|
| Roaming cloud upload | Graph Beta UserConfiguration API | Benefactor Circle (paid) |
| Template format | HTML only | Word (.docx) + HTML |
| Variable source | Graph API only | Graph + AD + LDAP + files |
| Transport rules | Built-in | Not included |
| Platform | Server-side / headless | Client-side (runs on user's machine) |
| Outlook add-in | Not included | Included (Benefactor Circle) |
| Cost | Free / MIT | Free core + paid Benefactor Circle |

For production environments managing thousands of users with complex assignment rules, **Set-OutlookSignatures with Benefactor Circle** is the recommended solution. This tool is designed for simpler deployments or as a starting point for custom integrations.

## References

- [Graph Beta UserConfiguration API (Jan 2026)](https://devblogs.microsoft.com/microsoft365dev/introducing-the-microsoft-graph-user-configuration-api-preview/)
- [Migrating EWS UserConfiguration to Graph](https://glenscales.substack.com/p/migrating-ews-getuserconfiguration)
- [Set-OutlookSignatures](https://github.com/Set-OutlookSignatures/Set-OutlookSignatures)
- [Frank Carius — outlookcloudsettings substrate API](https://www.msxfaq.de/cloud/exchangeonline/betrieb/outlookcloudsettings.htm)
- [Exchange Online InvokeCommand REST API](https://hajekj.net/2025/06/20/working-with-exchange-online-distribution-groups-via-rest/)

## License

MIT
