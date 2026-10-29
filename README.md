# PnP Modern Search – Additional data sources

**English** | [Deutsch](README.de.md)

An extensibility library for [PnP Modern Search](https://github.com/microsoft-search/pnp-modern-search) v4 that adds four data sources and a grouped file type filter to the Search Results and Search Filters web parts.

| Component | Purpose |
| --- | --- |
| **REST document search** | Document search via `/_api/search/postquery`; refiners computed over the whole result set (no 5,000 item threshold) |
| **OData** | Microsoft Graph, SharePoint REST v1/v2, Entra ID secured APIs and anonymous APIs |
| **SharePoint CAML** | Lists and libraries via `RenderListDataAsStream` with your own CAML |
| **Static data** | Items from a JSON object |
| **File type filter** | Filter control with file icons; extensions grouped by application (`xls`, `xlsx` → Excel) |

The user interface is available in English, German (`de-de`, `de-ch`, `de-at`), French, Spanish, Italian, Dutch, Portuguese (Brazil), Danish, Swedish, Norwegian (Bokmål), Finnish, Polish and Czech – the same languages as PnP Modern Search. English is used for all other languages. This documentation is available in English and German.

## Compatibility

| | Version |
| --- | --- |
| PnP Modern Search | 4.25.0 |
| SharePoint Framework | 1.23.0 (Heft toolchain) |
| `@pnp/modern-search-extensibility` | 2.1.0 |
| Node.js (build) | 22 (>= 22.14.0) |

The library must use the same SPFx version as the installed PnP Modern Search release ([compatibility matrix](https://microsoft-search.github.io/pnp-modern-search/extensibility/extensibility_compatibility_matrix/)).

## Installation

1. Download `pnp-modern-search-datasources.sppkg` from the [latest release](../../releases/latest) (or build it, see below).
2. Upload the package to the tenant app catalog and select **Make this solution available to all sites**.
3. Edit the page, open the **Search Results** web part and add the library ID in the extensibility configuration:

   ```
   8a686c92-47a4-49ac-ace4-74c323635408
   ```

4. For the file type filter, add the same ID to the **Search Filters** web part.
5. Reload the page. The new data sources and the filter template are now available.

Microsoft Graph and Entra ID secured APIs need API permissions on tenant level, for example:

```bash
m365 spo serviceprincipal grant add --resource 'Microsoft Graph' --scope 'Sites.Read.All'
m365 spo serviceprincipal grant add --resource '<app name>' --scope 'user_impersonation'
```

## Configuration

### Common settings (OData, SharePoint CAML, Static data)

| Setting | Description |
| --- | --- |
| Default sort | e.g. `Modified:desc,Title:asc` |
| Sortable fields for users | Fields offered in the sort control of the layout |
| Do not fetch data on page load | Query only when a search term or filter is set |
| Cache first result page | `localStorage` cache per web part instance, duration in minutes, "Clear cache" button |
| Fields for template slots | Title, link, summary, date, author, file type, ID; paths such as `fields.Title` or `Editor[0].title` |

### Tokens

| Token | Value |
| --- | --- |
| `{searchTerms}` / `{searchTermsOrEmpty}` | Search term (`*` or empty if none) |
| `{itemsCountPerPage}`, `{startRow}`, `{pageNumber}` | Paging |
| `{filter}` | OData filter built from the selected filters |
| `{sortField}`, `{sortDirection}` | Current sort |
| `{Site}`, `{SiteCollection}` | Current site / site collection URL |
| PnP tokens | `{User.Email}`, `{QueryString.x}`, `{Today}`, `{Page.<field>}` … |

Values are URL-encoded in URLs, JSON-escaped in request bodies and XML-escaped in CAML.

### REST document search

KQL query template (default `{searchTerms} IsDocument:1`), selected managed properties, default sort, result source ID, trim duplicates, query rules. Refiners are requested with `deephits=1000000`; date refiners get fixed intervals. Selected values are sent as FQL `RefinementFilters`.

### OData

| API type | URL formats | Authentication |
| --- | --- | --- |
| Microsoft Graph | `me/messages`, `/me`, `beta/me`, absolute URL | AadHttpClient (`https://graph.microsoft.com`) |
| SharePoint REST API | `_api/…` (current site), `/sites/x/_api/…`, absolute URL, `_api/v2.0/…` | SPHttpClient |
| Entra ID secured API | absolute URL + resource (`api://<app-id>`) | AadHttpClient |
| Anonymous | absolute URL | none |

- **Dynamic paging with server paging:** next pages from `@odata.nextLink`, discovered while paging. Example: `me/messages?$top={itemsCountPerPage}`
- **Dynamic paging without server paging:** paging via `{startRow}` / `{itemsCountPerPage}`.
- **Static paging:** load items once (optionally follow nextLinks up to a maximum), then page, filter and sort in the browser.
- **Sorting:** `$orderby` is appended unless the URL contains `{sortField}` or `$orderby`.
- **Filters:** client-side or server-side as OData `$filter` (combined with an existing `$filter` using `and`).

### SharePoint CAML

Site (empty = current), list by title, GUID or server-relative URL, view fields, recursive scope, CAML as `<View>`, `<Query>` or `<Where>` content. Search terms (`Contains` over configurable fields), selected filters, sorting and `RowLimit` are added automatically. Paging via `NextHref`. The extra field `AbsoluteFileRef` contains the absolute item URL.

### Static data

```json
{
  "items": [ { "linkName": "Microsoft Learn", "linkHref": "https://learn.microsoft.com" } ],
  "numberOfLinks": 1
}
```

Properties next to `items` are available in custom templates.

### File type filter

In the Search Filters web part, create a filter on the file type field (`FileType` for search, `File_x0020_Type` for CAML) and select the template **File types with icon (grouped)**. Clicking a group selects all of its extensions with OR. Works with every data source, including the built-in "SharePoint Search" and "Microsoft Search". Groups are defined in [`src/common/FileTypeGroups.ts`](src/common/FileTypeGroups.ts).

## Limitations

- List REST and CAML queries are subject to the 5,000 item list view threshold unless conditions and sorting use indexed columns.
- Client-side filters and counts only cover the loaded items.
- With nextLink and CAML paging the total count is unknown; pages are discovered while paging.
- External APIs must allow CORS for the SharePoint domain.

## Build from source

```bash
npm ci
npm run build
```

The package is written to `sharepoint/solution/pnp-modern-search-datasources.sppkg`.

### Add a language

1. Copy `src/loc/en-us.js` to `src/loc/<locale>.js` (e.g. `fr-fr.js`) and translate the values.
2. Build. SPFx picks up all files in `src/loc` automatically.

### Project structure

```
src/
  common/        Base class, client-side helpers, file type groups, tokens, cache
  components/    File type filter web component
  dataSources/   REST document search, OData, SharePoint CAML, static data
  libraries/     Extensibility library entry point (IExtensibilityLibrary)
  loc/           Localized strings (16 locales)
```

## Release

1. Update the version in `package.json` and `config/package-solution.json`, `CHANGELOG.md` and `RELEASE_NOTES.md`.
2. Commit, tag and push:

   ```bash
   git tag v2.3.2
   git push origin main --tags
   ```

3. The **Release** workflow builds the package and publishes a GitHub release with the `.sppkg` attached.

## Changelog

See [CHANGELOG.md](CHANGELOG.md).

## License

[MIT](LICENSE)
