# Changelog

All notable changes to this project are documented in this file.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses [Semantic Versioning](https://semver.org/).

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
