## PnP Modern Search – Additional data sources 2.3.1

Extensibility library for **PnP Modern Search 4.25.0** (SPFx 1.23.0) with four additional data sources and a grouped file type filter.

### What's new
- **Fix:** the package now passes app catalog validation (solution name with plain ASCII characters).
- **User interface in 16 locales**, the same languages as PnP Modern Search: Czech, Danish, German (`de-de`, `de-ch`, `de-at`), English, Spanish, Finnish, French, Italian, Norwegian (Bokmål), Dutch, Polish, Portuguese (Brazil), Swedish. English is the fallback.
- Code comments in English.

### Included
| Component | Description |
| --- | --- |
| REST document search | Search via `/_api/search/postquery`, refiners over the whole result set |
| OData | Microsoft Graph, SharePoint REST, Entra ID secured and anonymous APIs |
| SharePoint CAML | Lists and libraries via `RenderListDataAsStream` |
| Static data | Items from JSON |
| File type filter | File icons, extensions grouped by application |

### Installation
1. Download `pnp-modern-search-datasources.sppkg` below.
2. Upload it to the tenant app catalog and make it available to all sites.
3. Add the library ID `8a686c92-47a4-49ac-ace4-74c323635408` in the extensibility configuration of the Search Results and Search Filters web parts.

### Upgrade notes
- Upgrading from 2.0.x or earlier: data sources must be selected again (new internal keys since 2.1.0).

See [CHANGELOG.md](CHANGELOG.md) for all changes.
