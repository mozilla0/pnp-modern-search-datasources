# Changelog

All notable changes to this project are documented in this file.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses [Semantic Versioning](https://semver.org/).

## [2.3.4] - 2026-10-29

### Fixed
- REST document search: the property pane could not be opened after selecting the data source. `dataSourceProperties.selectedProperties` is now an array, as expected by the PnP web part (template slots, columns). Configurations stored by earlier versions are migrated automatically.

### Changed
- The managed properties of the REST document search are offered in the template slot field picker.

## [2.3.3] - 2026-10-29

### Fixed
- SharePoint CAML: a query that contains only the content of `<Where>` is now wrapped correctly (previously the condition was ignored).

### Added
- README (English and German): examples for every data source and the file type filter.

## [2.3.2] - 2026-10-29

### Fixed
- `Uncaught TypeError: t.some is not a function` in the result layouts: `dataSourceProperties.sortList` is now an array of sort field configurations, as expected by the built-in PnP sort control. Configurations stored by earlier versions are migrated automatically.

### Added
- The sort control of the PnP layouts now shows the "sortable fields for users" of all data sources, including the REST document search.

## [2.3.1] - 2026-10-29

### Fixed
- App catalog validation error ("The 'Name' attribute is invalid … Pattern constraint failed"): solution and feature names now use plain ASCII characters only.

### Changed
- Solution metadata (name, descriptions, developer website) in English.

## [2.3.0] - 2026-10-29

### Added
- User interface in the same languages as PnP Modern Search: `cs-cz`, `da-dk`, `de-de`, `en-us`, `es-es`, `fi-fi`, `fr-fr`, `it-it`, `nb-no`, `nl-nl`, `pl-pl`, `pt-br`, `sv-se` (plus `de-ch`, `de-at`).

### Changed
- Code comments translated to English.

## [2.2.0] - 2026-10-29

### Added
- Localization: all UI texts (data source names, property pane, filter control, error messages) in English and German (`en-us`, `de-de`, `de-ch`, `de-at`). English is the fallback for all other languages.
- GitHub workflows for build (push / pull request) and release (tag `v*`).
- English and German README, changelog and release notes.
- MIT license.

## [2.1.0] - 2026-10-29

### Added
- Filter control "File types with icon (grouped)": Fluent UI file icons, extensions grouped by application (e.g. `xls`, `xlsx` → Excel), selection of a group with OR, summed counts.

### Changed
- Neutral naming: package and solution `pnp-modern-search-datasources`, new internal data source keys (`Ext*`). Data sources configured with an earlier version must be selected again.

## [2.0.0] - 2026-10-29

### Added
- Data source "OData": Microsoft Graph, SharePoint REST v1/v2, Entra ID secured APIs, anonymous APIs; dynamic paging (nextLink or tokens) or static paging; `$orderby` / `$filter`.
- Data source "SharePoint CAML" via `RenderListDataAsStream` with tokens, search, filters, sorting and paging via `NextHref`.
- Data source "Static data" from JSON.
- Common settings: default sort, user sort fields, first page cache, "do not fetch on page load", template slot fields, PnP tokens.

### Removed
- Data source "Generic REST API" (replaced by "OData").

## [1.0.1] - 2026-10-29

### Fixed
- REST document search: refiners with `deephits=1000000` and date intervals.

## [1.0.0] - 2026-10-29

### Added
- Data source "REST document search" via `/_api/search/postquery` with server-side refiners.
