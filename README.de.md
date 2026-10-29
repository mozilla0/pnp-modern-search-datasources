# PnP Modern Search – Zusätzliche Datenquellen

[English](README.md) | **Deutsch**

Eine Erweiterungsbibliothek für [PnP Modern Search](https://github.com/microsoft-search/pnp-modern-search) v4. Sie ergänzt die Web Parts «Search Results» und «Search Filters» um vier Datenquellen und einen gruppierten Dateityp-Filter.

| Komponente | Zweck |
| --- | --- |
| **REST-Dokumentsuche** | Dokumentsuche über `/_api/search/postquery`; Verfeinerer über das gesamte Resultat (kein 5000er-Schwellenwert) |
| **OData** | Microsoft Graph, SharePoint REST v1/v2, Entra-ID-geschützte und anonyme APIs |
| **SharePoint CAML** | Listen und Bibliotheken über `RenderListDataAsStream` mit eigenem CAML |
| **Statische Daten** | Elemente aus einem JSON-Objekt |
| **Dateityp-Filter** | Filter-Steuerelement mit Dateisymbolen; Endungen nach Programm gruppiert (`xls`, `xlsx` → Excel) |

Die Benutzeroberfläche gibt es auf Englisch, Deutsch (`de-de`, `de-ch`, `de-at`), Französisch, Spanisch, Italienisch, Niederländisch, Portugiesisch (Brasilien), Dänisch, Schwedisch, Norwegisch (Bokmål), Finnisch, Polnisch und Tschechisch – dieselben Sprachen wie PnP Modern Search. Für alle anderen Sprachen wird Englisch verwendet. Diese Dokumentation gibt es auf Englisch und Deutsch.

## Kompatibilität

| | Version |
| --- | --- |
| PnP Modern Search | 4.25.0 |
| SharePoint Framework | 1.23.0 (Heft-Toolchain) |
| `@pnp/modern-search-extensibility` | 2.1.0 |
| Node.js (Build) | 22 (>= 22.14.0) |

## Installation

1. `pnp-modern-search-datasources.sppkg` aus dem [neusten Release](../../releases/latest) herunterladen (oder selbst bauen, siehe unten).
2. Paket in den Tenant-App-Katalog hochladen und «Für alle Websites verfügbar machen» wählen.
3. Im Web Part **Search Results** in der Erweiterbarkeits-Konfiguration die Library-ID eintragen:

   ```
   8a686c92-47a4-49ac-ace4-74c323635408
   ```

4. Für den Dateityp-Filter dieselbe ID auch im Web Part **Search Filters** eintragen.
5. Seite neu laden. Die Datenquellen und die Filtervorlage stehen jetzt zur Verfügung.

Microsoft Graph und Entra-ID-geschützte APIs brauchen API-Berechtigungen auf Tenant-Ebene, z. B.:

```bash
m365 spo serviceprincipal grant add --resource 'Microsoft Graph' --scope 'Sites.Read.All'
m365 spo serviceprincipal grant add --resource '<App-Name>' --scope 'user_impersonation'
```

## Konfiguration

Die Einstellungen sind in der [englischen README](README.md#configuration) ausführlich beschrieben, Beispiele folgen weiter unten; im Eigenschaftenbereich erscheinen sie auf Deutsch. Kurzüberblick:

- **Gemeinsam (OData, CAML, statische Daten):** Standardsortierung, sortierbare Felder für Benutzer, «Daten nicht beim Laden abrufen», Cache der ersten Seite, Felder für Vorlagen-Slots.
- **Tokens:** `{searchTerms}`, `{itemsCountPerPage}`, `{startRow}`, `{pageNumber}`, `{filter}`, `{sortField}`, `{sortDirection}`, `{Site}`, `{SiteCollection}` und alle PnP-Tokens.
- **REST-Dokumentsuche:** KQL-Vorlage, verwaltete Eigenschaften, Ergebnisquelle; Verfeinerer mit `deephits=1000000`.
- **OData:** API-Typ, URL, Header, Body; Paging dynamisch (nextLink oder Tokens) oder statisch; Filter clientseitig oder als `$filter`.
- **SharePoint CAML:** Website, Liste, View Fields, rekursiv, CAML mit Tokens; Suche, Filter, Sortierung und `RowLimit` werden ergänzt.
- **Dateityp-Filter:** Filter auf `FileType` (Suche) bzw. `File_x0020_Type` (CAML) anlegen und die Vorlage «Dateitypen mit Symbol (gruppiert)» wählen.

## Beispiele

Die Einstellungen verwenden die Bezeichnungen aus dem Eigenschaftenbereich. Nicht aufgeführte Einstellungen behalten ihren Standardwert.

### REST-Dokumentsuche: Office-Dokumente und PDFs der aktuellen Websitesammlung

Web Part «Search Results», Datenquelle **REST-Dokumentsuche (SharePoint Search)**:

| Einstellung | Wert |
| --- | --- |
| Abfragevorlage (KQL) | `{searchTerms} IsDocument:1 Path:"{SiteCollection}" (FileType:docx OR FileType:xlsx OR FileType:pptx OR FileType:pdf)` |
| Standardsortierung | `LastModifiedTime:desc` |
| Sortierbare Felder für Benutzer | `LastModifiedTime,Title` |

Web Part «Search Filters» (mit den Ergebnissen verbunden):

| Filterfeld | Vorlage |
| --- | --- |
| `FileType` | Dateitypen mit Symbol (gruppiert) |
| `LastModifiedTime` | Datumsintervall |
| `DisplayAuthor` | Kontrollkästchen |

Die Zählwerte der Filter gelten für das gesamte Resultat, auch bei Bibliotheken mit weit mehr als 5000 Dokumenten.

### OData – SharePoint REST API: Elemente einer Liste

Datenquelle **OData (Graph, SharePoint REST, API)**:

| Einstellung | Wert |
| --- | --- |
| API-Typ | SharePoint REST API |
| URL | `_api/web/lists/getbytitle('Projekte')/items?$select=Id,Title,Status,Kunde,Modified,FileRef&$top=500` |
| HTTP-Header (JSON) | `{ "Accept": "application/json;odata=nometadata" }` |
| Feld mit den Elementen | `value` |
| Paging | Statisch (Elemente laden, im Browser blättern) |
| Standardsortierung | `Modified:desc` |
| Sortierbare Felder für Benutzer | `Title,Modified` |
| Slots: Titel / Link / Zusammenfassung / Datum | `Title` / `FileRef` / `Kunde` / `Modified` |

Filter auf `Status` und `Kunde` werden im Browser angewendet, inklusive Zählwerte.

### OData – Microsoft Graph: E-Mails des angemeldeten Benutzers

| Einstellung | Wert |
| --- | --- |
| API-Typ | Microsoft Graph |
| URL | `me/messages?$select=subject,bodyPreview,webLink,receivedDateTime,from&$top={itemsCountPerPage}` |
| Feld mit den Elementen | `value` |
| Paging | Dynamisch (pro Seite eine Abfrage) |
| Server-Paging über nextLink verwenden | Ein |
| Sortierbare Felder für Benutzer | `receivedDateTime,subject` |
| Slots: Titel / Link / Zusammenfassung / Datum / Autor | `subject` / `webLink` / `bodyPreview` / `receivedDateTime` / `from.emailAddress.name` |

Nötige Berechtigung: `m365 spo serviceprincipal grant add --resource 'Microsoft Graph' --scope 'Mail.Read'`. Das gewählte Sortierfeld wird als `$orderby` übergeben.

### OData – Entra-ID-geschützte API (z. B. Azure Function)

| Einstellung | Wert |
| --- | --- |
| API-Typ | Entra-ID-geschützte API |
| URL | `https://func-contoso.azurewebsites.net/api/contracts?q={searchTermsOrEmpty}&skip={startRow}&top={itemsCountPerPage}&filter={filter}&orderby={sortField}%20{sortDirection}` |
| Entra-ID-Ressource | `api://11111111-2222-3333-4444-555555555555` |
| Feld mit den Elementen | `items` |
| Paging | Dynamisch (pro Seite eine Abfrage) |
| Server-Paging über nextLink verwenden | Aus |
| Feld mit der Gesamtanzahl | `total` |
| Filter anwenden | Serverseitig als OData-$filter |

Erwartete Antwort:

```json
{
  "total": 1287,
  "items": [
    { "id": "C-1001", "title": "Wartungsvertrag", "status": "Aktiv", "url": "https://…" }
  ]
}
```

Ein gewählter Filter auf `status` wird als `filter=status eq 'Aktiv'` übergeben. Die API muss CORS für die SharePoint-Domain erlauben, und die Berechtigung `user_impersonation` muss erteilt sein (siehe [Installation](#installation)).

### OData – anonyme API

Öffentlicher OData-Demodienst, Paging über Tokens und Filter auf dem Server:

| Einstellung | Wert |
| --- | --- |
| API-Typ | Anonym |
| URL | `https://services.odata.org/V4/TripPinService/People?$count=true&$top={itemsCountPerPage}&$skip={startRow}` |
| Feld mit den Elementen | `value` |
| Paging | Dynamisch (pro Seite eine Abfrage) |
| Server-Paging über nextLink verwenden | Aus |
| Feld mit der Gesamtanzahl | `@odata.count` |
| Filter anwenden | Serverseitig als OData-$filter |
| Slots: Titel / Zusammenfassung / ID | `FirstName` / `UserName` / `UserName` |

### SharePoint CAML: Word- und PDF-Dateien der letzten 30 Tage

| Einstellung | Wert |
| --- | --- |
| Website | `/sites/dms` |
| Liste oder Bibliothek | `/sites/dms/Shared Documents` |
| In Ordnern rekursiv suchen | Ein |
| Felder für Suchbegriffe | `Title,FileLeafRef` |
| Standardsortierung | *(leer – es gilt das `OrderBy` aus dem CAML)* |

CAML:

```xml
<View>
  <Query>
    <Where>
      <And>
        <In>
          <FieldRef Name="File_x0020_Type" />
          <Values>
            <Value Type="Text">docx</Value>
            <Value Type="Text">pdf</Value>
          </Values>
        </In>
        <Geq>
          <FieldRef Name="Modified" />
          <Value Type="DateTime"><Today OffsetDays="-30" /></Value>
        </Geq>
      </And>
    </Where>
    <OrderBy>
      <FieldRef Name="Modified" Ascending="FALSE" />
    </OrderBy>
  </Query>
</View>
```

Für die Bibliothek am besten die serverrelative URL verwenden: Sie funktioniert unabhängig von der Sprache der Website. Ab 5000 Elementen müssen `Modified` und `File_x0020_Type` indizierte Spalten sein.

### SharePoint CAML: meine offenen Aufgaben

| Einstellung | Wert |
| --- | --- |
| Liste oder Bibliothek | `Aufgaben` |
| Felder (View Fields, interne Namen) | `ID,Title,Status,DueDate,AssignedTo` |
| Slots: Titel / Datum / Autor | `Title` / `DueDate` / `AssignedTo[0].title` |

CAML (es genügt der Inhalt von `Where`):

```xml
<And>
  <Eq>
    <FieldRef Name="AssignedTo" LookupId="TRUE" />
    <Value Type="Integer"><UserID /></Value>
  </Eq>
  <Neq>
    <FieldRef Name="Status" />
    <Value Type="Text">Erledigt</Value>
  </Neq>
</And>
```

### Statische Daten: Schnelllinks mit Kategorien

Datenquelle **Statische Daten**:

```json
{
  "items": [
    { "linkName": "Microsoft 365 Admin Center", "linkHref": "https://admin.microsoft.com", "category": "Administration" },
    { "linkName": "SharePoint Admin Center", "linkHref": "https://admin.microsoft.com/sharepoint", "category": "Administration" },
    { "linkName": "Microsoft Learn", "linkHref": "https://learn.microsoft.com", "category": "Dokumentation" },
    { "linkName": "PnP Modern Search", "linkHref": "https://microsoft-search.github.io/pnp-modern-search/", "category": "Dokumentation" }
  ],
  "numberOfLinks": 4
}
```

| Einstellung | Wert |
| --- | --- |
| Standardsortierung | `linkName:asc` |
| Slots: Titel / Link / Zusammenfassung | `linkName` / `linkHref` / `category` |

Ein Web Part «Search Filters» mit dem Filterfeld `category` ergänzt einen Kategorienfilter. `numberOfLinks` steht in eigenen Vorlagen als `{{numberOfLinks}}` zur Verfügung.

### Dateityp-Filter mit beliebiger Datenquelle

Im Web Part «Search Filters» die Library-ID in der Erweiterbarkeits-Konfiguration eintragen und einen Filter anlegen:

| Datenquelle | Filterfeld | Vorlage |
| --- | --- | --- |
| SharePoint Search, Microsoft Search, REST-Dokumentsuche | `FileType` | Dateitypen mit Symbol (gruppiert) |
| SharePoint CAML | `File_x0020_Type` | Dateitypen mit Symbol (gruppiert) |
| OData / Statische Daten | Feld mit der Dateiendung, z. B. `fileType` | Dateitypen mit Symbol (gruppiert) |

## Grenzen

- Listen-REST- und CAML-Abfragen unterliegen dem Schwellenwert von 5000 Elementen, sofern Bedingungen und Sortierung nicht auf indizierten Spalten liegen.
- Clientseitige Filter und Zählwerte beziehen sich nur auf die geladenen Elemente.
- Bei nextLink- und CAML-Paging ist die Gesamtanzahl unbekannt.
- Fremde APIs müssen CORS für die SharePoint-Domain erlauben.

## Selbst bauen

```bash
npm ci
npm run build
```

Das Paket liegt danach unter `sharepoint/solution/pnp-modern-search-datasources.sppkg`.

Weitere Sprache: `src/loc/en-us.js` nach `src/loc/<locale>.js` kopieren, übersetzen und neu bauen.

## Release

1. Version in `package.json` und `config/package-solution.json` sowie `CHANGELOG.md` und `RELEASE_NOTES.md` nachführen.
2. Commit, Tag und Push:

   ```bash
   git tag v2.3.3
   git push origin main --tags
   ```

3. Der Workflow **Release** baut das Paket und veröffentlicht ein GitHub-Release mit angehängtem `.sppkg`.

## Änderungen

Siehe [CHANGELOG.md](CHANGELOG.md).

## Lizenz

[MIT](LICENSE)
