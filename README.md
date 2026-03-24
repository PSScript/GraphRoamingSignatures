# Graph Roaming Signatures

E-Mail-Signaturverwaltung für Microsoft 365 / Exchange Online über die Microsoft Graph API.

> **Schnellstart für den täglichen Betrieb:** Siehe [Schnellstart.md](Schnellstart.md) — praxisnahe Anleitung für L1-Support, Azubis und alle, die Signatur-Tickets bearbeiten.

---

## Überblick

Dieses Projekt deployt HTML-E-Mail-Signaturen automatisiert in Exchange Online Postfächer.
Es nutzt die **Graph Beta UserConfiguration API** (`MailboxConfigItem.ReadWrite`), die Microsoft
im Januar 2026 veröffentlicht hat — nicht das veraltete `Set-MailboxMessageConfiguration`,
das seit der Einführung von Roaming Signatures nicht mehr zuverlässig funktioniert.

### Kernfunktionen

- **Multi-Company:** Unterschiedliche Vorlagen pro Tochterfirma, gesteuert über das Entra ID Feld `companyName`
- **VIP-Overrides:** GFs/CEOs bekommen eigene Signaturen mit Foto, LinkedIn-Link, individueller Gestaltung
- **Gruppen-Overrides:** Temporäre Kampagnen-Signaturen über Security Groups zuweisen
- **Kampagnenbanner:** Optional unter jeder Signatur (480×96, 360×56 oder frei definierbar)
- **Delta-Erkennung:** Graph `/users/delta` + SHA256-Prüfsumme in `extensionAttribute` — nur geänderte Benutzer werden versorgt
- **Encoding-Handling:** Windows-1252 / ISO-8859-1 / UTF-8 / UTF-16 automatische Erkennung und Konvertierung

### Zuordnungs-Hierarchie

```
VIP-Override (UPN)  →  Security Group  →  companyName  →  Standard-Vorlage
     (1. Priorität)        (2.)              (3.)            (Fallback)
```

Erste Übereinstimmung gewinnt.

---

## Dateien

| Datei | Beschreibung |
|---|---|
| `Set-GraphSignature.ps1` | **Hauptskript** — alles in einer Datei, Konfiguration inline oben im Skript |
| `Invoke-SignatureAutomation.ps1` | Alternative mit externer `config.json` für komplexere Setups |
| `Manage-OutlookSignatures.ps1` | Multi-Strategie-Engine (Roaming + OWA-Fallback + Transport Rules) |
| `Register-SignatureManagerApp.ps1` | Entra ID App-Registrierung mit allen nötigen Berechtigungen |
| `config.json` | Beispiel-Konfiguration für `Invoke-SignatureAutomation.ps1` |
| `templates/corporate.htm` | Beispiel-HTML-Vorlage |
| `designer/SignatureDesigner.jsx` | React-basierter WYSIWYG-Editor für Signatur-Vorlagen |
| `Schnellstart.md` | Deutschsprachige Anleitung für den täglichen Betrieb |

### Welches Skript nehmen?

- **`Set-GraphSignature.ps1`** — für die meisten Umgebungen. Alles in einer Datei, keine externen Abhängigkeiten. Konfiguration direkt im Skript oben editieren.
- **`Invoke-SignatureAutomation.ps1`** — wenn die Konfiguration in einer separaten JSON-Datei liegen soll (z.B. weil mehrere Admins daran arbeiten oder die Config in Git versioniert wird).
- **`Manage-OutlookSignatures.ps1`** — wenn zusätzlich OWA-Fallback oder Transport Rules als Deployment-Strategie gebraucht werden.

---

## Voraussetzungen

### Entra ID App-Registrierung

Einmal ausführen:

```powershell
.\Register-SignatureManagerApp.ps1 -AppName "Signature Manager" -CreateClientSecret
```

Oder manuell im [Entra Admin Center](https://entra.microsoft.com) anlegen.

### Benötigte Berechtigungen

| Berechtigung | Typ | Wofür |
|---|---|---|
| `MailboxConfigItem.ReadWrite` | Delegiert + App | Roaming Signatures über UserConfiguration API |
| `Mail.ReadWrite` | Delegiert + App | Postfachzugriff |
| `User.Read.All` | Delegiert + App | Benutzerdaten für Variablen-Ersetzung |
| `MailboxSettings.ReadWrite` | Delegiert + App | Abwesenheitsnachrichten |
| `Exchange.ManageAsApp` | Application | Transport Rules, EXO InvokeCommand |

### PowerShell

- PowerShell 7.x+ empfohlen (5.1 funktioniert für die meisten Features)
- Keine externen Module nötig — das Skript arbeitet direkt mit REST-Aufrufen
- `Microsoft.Graph.Applications` nur für `Register-SignatureManagerApp.ps1` benötigt

---

## Schnellstart

### 1. App registrieren

```powershell
.\Register-SignatureManagerApp.ps1 -AppName "Signature Manager" -CreateClientSecret
```

Die ausgegebene `ClientId` und `ClientSecret` notieren.

### 2. Konfiguration anpassen

In `Set-GraphSignature.ps1` die ersten Zeilen des Konfigurationsblocks anpassen:

```powershell
$TenantId     = 'eure-tenant-id'
$ClientId     = 'eure-client-id'
$ClientSecret = 'euer-secret'
$PrimaryDomain = 'eure-domain.de'
```

Firmen-Zuordnung, VIP-Overrides und Kampagnenbanner nach Bedarf konfigurieren.

### 3. Testlauf

```powershell
# Vorschau — ändert nichts
.\Set-GraphSignature.ps1 -WhatIf

# Einzelner Benutzer
.\Set-GraphSignature.ps1 -UserUPN "test.user@eure-domain.de"

# Alle Benutzer
.\Set-GraphSignature.ps1
```

### 4. Geplante Ausführung einrichten

Task Scheduler, Azure Automation oder cron — täglich oder alle 4 Stunden.

---

## Variablen-Syntax

Vorlagen verwenden `%Variable%`-Platzhalter:

```html
<td>%DisplayName%</td>
<td>%JobTitle% · %Department%</td>
<td>%Phone%</td>
<td>%ExtensionAttribute7%</td>
```

Vollständige Variablenliste: Siehe [Schnellstart.md](Schnellstart.md#verfügbare-variablen).

---

## Layout

### Standard-Signatur

```
┌──────────────────────────────────────────────────────┐
│                                          │           │
│  Max Mustermann                          │  ┌─────┐  │
│  SENIOR CONSULTANT                       │  │Logo │  │
│                                          │  │oder │  │
│  T  +49 711 123456-0                     │  │Foto │  │
│  E  m.mustermann@contoso.com             │  └─────┘  │
│                                          │ 240×160px │
│  Contoso Group · Musterstr. 7 · 12345    │           │
└──────────────────────────────────────────────────────┘
```

- **Links:** Kontaktdaten
- **Rechts:** Logo (160×60) oder VIP-Foto (240×160)
- **Trennlinie:** 3px vertikal in Akzentfarbe

### Kampagnenbanner (optional)

| Größe | Verwendung |
|---|---|
| 480 × 96 px | Standard (Messen, Events) |
| 360 × 56 px | Kompakt (Dauerkampagnen) |
| 480 × 120 px | Groß (Produktlaunch) |

---

## Änderungserkennung

Das Skript ist für den täglichen Betrieb optimiert:

1. **Graph Delta Query** — fragt nur Benutzer ab, deren Daten sich seit dem letzten Lauf geändert haben
2. **Prüfsumme** — SHA256 über alle signaturrelevanten Felder + Template-Hash, gespeichert in `extensionAttribute15` als `SIG:a3f2c1b98d21|2026-03-23T14:00Z`
3. **Template-Änderung** — neues Template = neuer Hash = automatischer Re-Deploy aller betroffenen Benutzer

Täglicher Lauf bei 5000 Benutzern: unter 2 Minuten (nur 10–20 tatsächliche Deployments).

---

## Verwandte Projekte

- [Set-OutlookSignatures](https://github.com/Set-OutlookSignatures/Set-OutlookSignatures) — umfassende Client-seitige Lösung mit DOCX-Templates, Outlook-Add-In und Roaming Signature Sync (Benefactor Circle)
- [Graph Beta UserConfiguration API](https://devblogs.microsoft.com/microsoft365dev/introducing-the-microsoft-graph-user-configuration-api-preview/) — die offizielle Microsoft-Dokumentation (Januar 2026)
- [EWS → Graph Migration](https://glenscales.substack.com/p/migrating-ews-getuserconfiguration) — Hintergrundartikel zur UserConfiguration-Migration

---

## Lizenz

MIT
