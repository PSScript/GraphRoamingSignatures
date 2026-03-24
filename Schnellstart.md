# Schnellstart — Signaturverwaltung mit Graph Roaming Signatures

> Dieses Dokument ist für alle gedacht, die im Tagesgeschäft mit E-Mail-Signaturen arbeiten:
> L1-Support, Azubis, Admins — jeder, der Tickets wie *"Meine Signatur stimmt nicht"* oder
> *"Der neue GF braucht eine eigene Signatur"* bearbeiten muss.

---

## Wie funktioniert das Ganze?

Das Skript `Set-GraphSignature.ps1` macht folgendes:

1. Holt Benutzerdaten aus dem Entra ID (ehemals Azure AD) über die Microsoft Graph API
2. Setzt die Daten in eine HTML-Vorlage ein (Name, Telefon, Abteilung usw.)
3. Schreibt die fertige Signatur in das Exchange Online Postfach des Benutzers
4. Merkt sich per Prüfsumme, welche Benutzer schon aktuell sind — beim nächsten Lauf werden nur geänderte Benutzer neu versorgt

**Wichtig:** Die Signatur wird **serverseitig** gesetzt. Der Benutzer muss nichts tun.
Beim nächsten Öffnen von Outlook (New Outlook / OWA) ist die Signatur da.


---

## Die Konfiguration — alles in einer Datei

Alles Wichtige steht oben in `Set-GraphSignature.ps1` im Konfigurationsblock.
Man muss **keine** externen Dateien bearbeiten.

### Domain und Basis-URL

```powershell
$PrimaryDomain = 'contoso.com'                              # ← Eure Hauptdomain
$AssetBaseUrl  = "https://assets.$PrimaryDomain/signatures"  # ← Wo liegen Logos/Bilder?
```

Alles andere (Ausschlüsse, VIP-Adressen) leitet sich von `$PrimaryDomain` ab.
Domain einmal ändern → fertig.


### Firmen-Zuordnung

Jeder Benutzer hat in Entra ID ein Feld `companyName`. Das Skript schaut nach diesem Feld
und wählt die passende Vorlage:

```powershell
$CompanyTemplates = @{
    'Contoso Group' = @{
        SignatureName = 'Contoso Group Signature'
        LogoUrl       = "$AssetBaseUrl/logo-contoso.png"
        AccentColor   = '#E30613'                          # ← Farbe der seitlichen Linie
    }
    'Contoso Cloud' = @{
        SignatureName = 'Contoso Cloud Signature'
        LogoUrl       = "$AssetBaseUrl/logo-contoso-cloud.png"
        AccentColor   = '#E30613'
    }
}
```

**Neue Firma hinzufügen?** Einfach einen neuen Block reinkopieren:

```powershell
    'Neue Tochter GmbH' = @{
        SignatureName = 'Neue Tochter Signature'
        LogoUrl       = "$AssetBaseUrl/logo-neue-tochter.png"
        AccentColor   = '#336699'
    }
```

> Der Key (`'Neue Tochter GmbH'`) muss **exakt** dem `companyName` in Entra ID entsprechen.
> Groß-/Kleinschreibung ist egal, aber Leerzeichen und Sonderzeichen müssen stimmen.


---

## Variablen-Syntax — das `%Variable%`-System

In der HTML-Vorlage werden Platzhalter mit Prozentzeichen geschrieben:

```html
<td>%DisplayName%</td>
<td>%JobTitle% · %Department%</td>
<td>%Phone%</td>
<td>%ExtensionAttribute7%</td>
```

Das Skript ersetzt diese Platzhalter mit den echten Daten aus Entra ID.

### Verfügbare Variablen

| Variable | Entra ID Feld | Beispiel |
|---|---|---|
| `%DisplayName%` | displayName | Max Mustermann |
| `%GivenName%` | givenName | Max |
| `%Surname%` | surname | Mustermann |
| `%JobTitle%` | jobTitle | Senior Consultant |
| `%Department%` | department | IT Infrastructure |
| `%Mail%` | mail | m.mustermann@contoso.com |
| `%Phone%` | businessPhones[0] | +49 711 123456-0 |
| `%Mobile%` | mobilePhone | +49 170 1234567 |
| `%Fax%` | faxNumber | +49 711 123456-99 |
| `%Company%` | companyName | Contoso Group |
| `%Office%` | officeLocation | Main Office |
| `%Street%` | streetAddress | Musterstraße 7 |
| `%City%` | city | Musterstadt |
| `%PostalCode%` | postalCode | 12345 |
| `%State%` | state | Bundesland |
| `%Country%` | country | Germany |
| `%MailNickname%` | mailNickname | m.mustermann |
| `%ManagerName%` | manager → displayName | Erika Chefin |
| `%ManagerMail%` | manager → mail | e.chefin@contoso.com |
| `%ExtensionAttribute1%` bis `%ExtensionAttribute15%` | onPremisesExtensionAttributes | (frei belegbar) |

### Leere Felder

Wenn ein Feld in Entra ID leer ist (z.B. kein Fax), wird die Variable durch einen
leeren String ersetzt. **Komplett leere Tabellenzeilen werden automatisch entfernt** —
es entstehen keine hässlichen Lücken.

### ExtensionAttributes — die Geheimwaffe

Die Felder `extensionAttribute1` bis `extensionAttribute15` sind frei belegbar.
Typische Verwendung:

| Attribut | Verwendung (Beispiel) |
|---|---|
| `extensionAttribute1` | Standortkürzel (STR, HH, BER) |
| `extensionAttribute7` | Rechtlicher Hinweis / Disclaimer |
| `extensionAttribute10` | Persönliche Zusatzzeile ("Certified Azure Expert") |
| `extensionAttribute14` | Kampagnen-Flag (wird im Skript nicht direkt genutzt) |
| `extensionAttribute15` | **Reserviert für Signatur-Prüfsumme** — nicht manuell ändern! |

> **Achtung:** `extensionAttribute15` wird vom Skript als Speicher für die Änderungserkennung
> verwendet. Dort steht z.B. `SIG:a3f2c1b98d21|2026-03-23T14:00Z`.
> **Dieses Feld niemals manuell überschreiben**, sonst wird der Benutzer beim nächsten Lauf
> unnötig neu versorgt (schadet nicht, kostet aber Zeit).


---

## Das HTML-Layout

### Standardlayout

```
┌──────────────────────────────────────────────────────┐
│                                          │           │
│  Max Mustermann                          │  ┌─────┐  │
│  SENIOR CONSULTANT                       │  │     │  │
│                                          │  │Logo │  │
│  T  +49 711 123456-0                     │  │oder │  │
│  M  +49 170 1234567                      │  │Foto │  │
│  E  m.mustermann@contoso.com             │  │     │  │
│                                          │  └─────┘  │
│  Contoso Group · Musterstr. 7 · 12345    │ 240×160px │
│                                          │           │
└──────────────────────────────────────────────────────┘
```

- **Links:** Kontaktdaten (Text, linksbündig)
- **Rechts:** Logo (160×60px) oder bei VIPs: Portraitfoto (240×160px)
- **Trennlinie:** 3px vertikale Linie in der Akzentfarbe

### Kampagnenbanner

Optional kann unter der Signatur ein Banner angehängt werden:

```
┌──────────────────────────────────────────────────────┐
│  [normale Signatur wie oben]                         │
├──────────────────────────────────────────────────────┤
│  ┌──────────────────────────────────────────────┐    │
│  │         Kampagnenbanner 480×96px              │    │
│  └──────────────────────────────────────────────┘    │
└──────────────────────────────────────────────────────┘
```

Empfohlene Bannergrößen:

| Größe | Verwendung |
|---|---|
| **480 × 96 px** | Standard-Kampagnenbanner (Messen, Events) |
| **360 × 56 px** | Kompaktes Banner (dezenter, für Dauerkampagnen) |
| **480 × 120 px** | Großes Banner (Produktlaunch, wichtige Ankündigung) |

Banner aktivieren in der Konfiguration:

```powershell
$CampaignBanner = @{
    ImageUrl = "$AssetBaseUrl/banners/kampagne-q2.png"
    Width    = 480
    Height   = 96
    AltText  = 'Besuchen Sie uns auf der IT-SA 2026'
    LinkUrl  = "https://www.$PrimaryDomain/events/itsa"
}
```

Banner deaktivieren: `$CampaignBanner = $null`


---

## Die häufigsten Aufgaben

### 1. "Benutzer hat falsche Signatur"

**Prüfe zuerst im Entra ID**, ob die Daten stimmen:

```powershell
# Im Graph Explorer oder per PowerShell:
Get-MgUser -UserId "user@contoso.com" -Property displayName,jobTitle,department,businessPhones,companyName
```

Wenn die Daten in Entra ID falsch sind → dort korrigieren. Beim nächsten Skript-Lauf
wird die Signatur automatisch aktualisiert (Prüfsumme ändert sich).

Wenn die Daten stimmen, aber die Signatur trotzdem alt ist → Skript manuell für
diesen Benutzer anstoßen:

```powershell
.\Set-GraphSignature.ps1 -UserUPN "user@contoso.com" -Force
```

### 2. "Neuer Mitarbeiter braucht eine Signatur"

**Nichts tun.** Sobald der Benutzer in Entra ID angelegt ist und ein `companyName` hat,
wird er beim nächsten geplanten Skript-Lauf automatisch versorgt.

Wenn es schnell gehen muss:

```powershell
.\Set-GraphSignature.ps1 -UserUPN "neuer.mitarbeiter@contoso.com"
```

### 3. "GF/CEO will eine besondere Signatur"

VIP-Override in der Konfiguration hinzufügen:

```powershell
$VipOverrides = @{
    # ... bestehende Einträge ...
    "neuer.gf@contoso.com" = @{
        SignatureName = 'GF Neue Niederlassung'
        LogoUrl       = "$AssetBaseUrl/logo-contoso.png"
        PhotoUrl      = "$AssetBaseUrl/photos/neuer-gf.jpg"   # ← Portraitfoto 240×160
        AccentColor   = '#E30613'
        LinkedIn      = 'https://linkedin.com/in/neuer-gf'    # ← optional
    }
}
```

Dann einmal ausführen:

```powershell
.\Set-GraphSignature.ps1 -UserUPN "neuer.gf@contoso.com" -Force
```

### 4. "Neue Tochterfirma / Marke hinzufügen"

1. Logo als PNG bereitstellen unter `$AssetBaseUrl/logo-neuefirma.png`
2. Eintrag in `$CompanyTemplates` hinzufügen:

```powershell
    'Neue Firma GmbH' = @{
        SignatureName = 'Neue Firma Signature'
        LogoUrl       = "$AssetBaseUrl/logo-neuefirma.png"
        AccentColor   = '#336699'
    }
```

3. Sicherstellen, dass die Benutzer in Entra ID `companyName = "Neue Firma GmbH"` haben
4. Skript mit `-Force` laufen lassen (da neue Firma = alle Benutzer dieser Firma betrifft)

### 5. "Kampagnenbanner für alle aktivieren"

```powershell
# In der Konfiguration $CampaignBanner setzen:
$CampaignBanner = @{
    ImageUrl = "$AssetBaseUrl/banners/mein-banner.png"
    Width    = 480
    Height   = 96
    AltText  = 'Unser Event 2026'
    LinkUrl  = 'https://www.contoso.com/event'
}
```

Dann: `.\Set-GraphSignature.ps1 -Force` (alle Benutzer neu versorgen).

**Banner wieder entfernen:** `$CampaignBanner = $null` setzen und erneut `-Force`.

### 6. "Nur prüfen, was passieren würde"

```powershell
.\Set-GraphSignature.ps1 -WhatIf
```

Zeigt an, welche Benutzer betroffen wären, ohne etwas zu ändern.

### 7. "Alle Signaturen komplett neu ausrollen" (Rebrand, Template-Änderung)

```powershell
.\Set-GraphSignature.ps1 -Force
```

Ignoriert alle Prüfsummen. Jeder Benutzer bekommt seine Signatur neu.
Bei 5000 Benutzern ca. 30–45 Minuten (wegen API-Throttling).


---

## Zuordnungs-Hierarchie — wer bekommt was?

Das Skript prüft für jeden Benutzer in dieser Reihenfolge:

```
1. VIP-Override     Hat der Benutzer einen Eintrag in $VipOverrides?
       ↓ nein           → Falls ja: VIP-Vorlage verwenden (Foto, LinkedIn usw.)
2. Gruppen-Override Ist der Benutzer Mitglied einer Gruppe in $GroupOverrides?
       ↓ nein           → Falls ja: Gruppen-Vorlage verwenden (z.B. Recruiting-Banner)
3. Firmenname       Passt sein companyName zu einem Key in $CompanyTemplates?
       ↓ nein           → Falls ja: Firmen-Vorlage verwenden (eigenes Logo, Farbe)
4. Standard         → Standard-Vorlage ($DefaultTemplate) verwenden
```

**Erste Übereinstimmung gewinnt.** Ein CEO bekommt immer seine VIP-Vorlage,
auch wenn er gleichzeitig in einer Marketing-Gruppe ist.


---

## Änderungserkennung — warum das Skript so schnell ist

Bei jedem Lauf passiert folgendes:

1. **Graph Delta Query** fragt Microsoft: *"Welche Benutzer haben sich seit meinem
   letzten Aufruf geändert?"* — statt 5000 Benutzer kommen nur die 50 zurück,
   bei denen sich etwas getan hat.

2. Für jeden dieser 50 Benutzer wird eine **Prüfsumme** berechnet:
   `SHA256(Name + Titel + Telefon + Abteilung + ... + Template-Hash)`

3. Diese Prüfsumme wird mit dem Wert in `extensionAttribute15` verglichen:
   - **Stimmt überein?** → Überspringen (nichts hat sich geändert)
   - **Stimmt nicht?** → Signatur neu deployen und neue Prüfsumme speichern

4. Wenn sich eine **Vorlage** ändert (HTML-Datei bearbeitet), ändert sich der
   Template-Hash → **alle** Prüfsummen passen nicht mehr → alle Benutzer dieser
   Vorlage werden automatisch neu versorgt.

**Ergebnis:** Ein täglicher Lauf über 5000 Benutzer dauert typischerweise unter 2 Minuten,
weil nur 10–20 Benutzer tatsächlich versorgt werden müssen.


---

## Fehlerbehebung

### "Cannot write checksum — attribute synced from on-prem AD"

Das `extensionAttribute15` wird von AD Connect aus dem lokalen Active Directory synchronisiert.
Das Skript kann den Wert nicht nach Entra ID zurückschreiben.

**Lösung:** Entweder ein anderes Attribut verwenden, das nicht synchronisiert wird
(`$ChecksumAttribute = 'extensionAttribute14'`), oder das Attribut im AD Connect
aus der Synchronisation ausschließen.

### "Throttled 429 — waiting 30s"

Microsoft drosselt die API-Aufrufe. Das ist normal und kein Fehler.
Das Skript wartet automatisch und versucht es erneut. Bei sehr großen Läufen
(>5000 Benutzer) den `$ThrottleDelayMs`-Wert erhöhen.

### "Signatur wird nicht angezeigt in Outlook Desktop (klassisch)"

Das Skript setzt die Signatur über die **UserConfiguration API**, die von
OWA und New Outlook gelesen wird. Outlook Desktop (klassisch, Win32) liest
seine Signaturen aus lokalen Dateien unter `%APPDATA%\Microsoft\Signatures`.

Für Outlook Desktop braucht man entweder:
- [Set-OutlookSignatures](https://github.com/Set-OutlookSignatures/Set-OutlookSignatures)
  (Client-seitige Lösung, setzt lokale Dateien)
- Roaming Signatures aktiviert (`DisableRoamingSignatures = 0` in der Registry)

### "Benutzer hat gar keine Signatur"

Checkliste:
1. Hat der Benutzer ein `mail`-Feld in Entra ID? (ohne Mail → wird übersprungen)
2. Ist die UPN in `$ExcludeUPNs`? Oder beginnt mit einem Prefix aus `$ExcludeUPNPrefixes`?
3. Ist der `companyName` in `$ExcludeCompanyNames`?
4. Passt der `companyName` zu keinem Eintrag und es gibt kein `$DefaultTemplate`?

Prüfen mit:

```powershell
.\Set-GraphSignature.ps1 -UserUPN "user@contoso.com" -WhatIf
```


---

## Geplante Ausführung (Scheduled Task / Azure Automation)

### Windows Task Scheduler

```
Programm:   pwsh.exe
Argumente:  -NonInteractive -NoProfile -File "C:\Scripts\Set-GraphSignature.ps1"
Zeitplan:   Täglich, 06:00 Uhr
Ausführen als: Dienstkonto mit Netzwerkzugriff
```

### Azure Automation Runbook

1. `Set-GraphSignature.ps1` als Runbook hochladen
2. Zeitplan erstellen (täglich oder alle 4 Stunden)
3. `$ClientId`, `$ClientSecret`, `$TenantId` als verschlüsselte Variablen hinterlegen
4. `$StatePath` auf einen Azure Blob Storage oder Automation-Variable umstellen


---

## Kurzreferenz

| Was will ich? | Befehl |
|---|---|
| Einzelnen Benutzer aktualisieren | `.\Set-GraphSignature.ps1 -UserUPN "user@contoso.com"` |
| Einzelnen Benutzer erzwingen | `.\Set-GraphSignature.ps1 -UserUPN "user@contoso.com" -Force` |
| Nur schauen, nicht anfassen | `.\Set-GraphSignature.ps1 -WhatIf` |
| Alle Benutzer komplett neu | `.\Set-GraphSignature.ps1 -Force` |
| Normaler geplanter Lauf | `.\Set-GraphSignature.ps1` |

| Was will ich ändern? | Wo? |
|---|---|
| Domain | `$PrimaryDomain` |
| Neue Firma hinzufügen | `$CompanyTemplates` — neuen Block einfügen |
| VIP-Signatur anlegen | `$VipOverrides` — UPN als Key |
| Kampagnenbanner | `$CampaignBanner` setzen oder `$null` |
| Akzentfarbe einer Firma | `AccentColor` im jeweiligen `$CompanyTemplates`-Eintrag |
| Logo tauschen | `LogoUrl` im jeweiligen Eintrag ändern |
| Prüfsummen-Attribut | `$ChecksumAttribute` (Standard: `extensionAttribute15`) |
| Benutzer ausschließen | `$ExcludeUPNs` oder `$ExcludeUPNPrefixes` |

---

*Letzte Aktualisierung: März 2026*
