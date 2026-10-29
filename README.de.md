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

Die Einstellungen sind in der [englischen README](README.md#configuration) ausführlich beschrieben; im Eigenschaftenbereich erscheinen sie auf Deutsch. Kurzüberblick:

- **Gemeinsam (OData, CAML, statische Daten):** Standardsortierung, sortierbare Felder für Benutzer, «Daten nicht beim Laden abrufen», Cache der ersten Seite, Felder für Vorlagen-Slots.
- **Tokens:** `{searchTerms}`, `{itemsCountPerPage}`, `{startRow}`, `{pageNumber}`, `{filter}`, `{sortField}`, `{sortDirection}`, `{Site}`, `{SiteCollection}` und alle PnP-Tokens.
- **REST-Dokumentsuche:** KQL-Vorlage, verwaltete Eigenschaften, Ergebnisquelle; Verfeinerer mit `deephits=1000000`.
- **OData:** API-Typ, URL, Header, Body; Paging dynamisch (nextLink oder Tokens) oder statisch; Filter clientseitig oder als `$filter`.
- **SharePoint CAML:** Website, Liste, View Fields, rekursiv, CAML mit Tokens; Suche, Filter, Sortierung und `RowLimit` werden ergänzt.
- **Dateityp-Filter:** Filter auf `FileType` (Suche) bzw. `File_x0020_Type` (CAML) anlegen und die Vorlage «Dateitypen mit Symbol (gruppiert)» wählen.

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
   git tag v2.3.2
   git push origin main --tags
   ```

3. Der Workflow **Release** baut das Paket und veröffentlicht ein GitHub-Release mit angehängtem `.sppkg`.

## Änderungen

Siehe [CHANGELOG.md](CHANGELOG.md).

## Lizenz

[MIT](LICENSE)
