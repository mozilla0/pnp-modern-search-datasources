import { SortFieldDirection } from '@pnp/modern-search-extensibility';

/**
 * Sort field configuration in the shape the PnP Modern Search layouts expect in
 * `dataSourceProperties.sortList` (rendered by the built-in `pnp-sortfield` control).
 */
export interface ISortFieldConfiguration {
  sortField: string;
  sortDirection: SortFieldDirection;
  isDefaultSort: boolean;
  isUserSort: boolean;
  sortFieldDisplayName: string;
}

function splitList(value: string): string[] {
  return (value || '').split(',').map(s => s.trim()).filter(s => !!s);
}

/**
 * Builds the PnP sort list from the text settings
 * "default sort" (e.g. "Modified:desc,Title:asc") and "user sort fields" (e.g. "Title,Modified").
 */
export function buildSortList(defaultSort: string, userSortFields: string): ISortFieldConfiguration[] {
  const userFields: string[] = splitList(userSortFields);
  const result: ISortFieldConfiguration[] = [];

  splitList(defaultSort).forEach(entry => {
    const parts: string[] = entry.split(':');
    const field: string = parts[0].trim();
    if (!field) { return; }
    result.push({
      sortField: field,
      sortDirection: (parts[1] || '').trim().toLowerCase() === 'desc' ? SortFieldDirection.Descending : SortFieldDirection.Ascending,
      isDefaultSort: true,
      isUserSort: userFields.indexOf(field) !== -1,
      sortFieldDisplayName: field
    });
  });

  userFields.forEach(field => {
    if (!result.some(r => r.sortField === field)) {
      result.push({
        sortField: field,
        sortDirection: SortFieldDirection.Ascending,
        isDefaultSort: false,
        isUserSort: true,
        sortFieldDisplayName: field
      });
    }
  });

  return result;
}
