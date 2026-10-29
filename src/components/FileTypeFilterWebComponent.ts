import { getFileTypeIconAsUrl } from '@fluentui/react-file-type-icons';
import {
  BaseWebComponent,
  ExtensibilityConstants,
  FilterComparisonOperator,
  FilterConditionOperator,
  IDataFilterInfo,
  IDataFilterInternal,
  IDataFilterValueInternal
} from '@pnp/modern-search-extensibility';
import { getFileTypeGroup, IFileTypeGroup } from '../common/FileTypeGroups';

interface IGroupedValue {
  group: IFileTypeGroup;
  members: IDataFilterValueInternal[];
  count: number;
  selected: boolean;
  disabled: boolean;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type ThemeVariant = { palette?: { [k: string]: string }; semanticColors?: { [k: string]: string } } | any;

export const FILE_TYPE_FILTER_COMPONENT: string = 'pnp-ext-filetype-filter';

/**
 * File type filter control: shows one icon per application (Word, Excel, PDF …),
 * groups the extensions (xls + xlsx = Excel) and, when clicked, selects all extensions
 * of the group combined with OR. Works with any data source that returns file
 * extensions as filter values (e.g. FileType, File_x0020_Type).
 */
export class FileTypeFilterWebComponent extends BaseWebComponent {

  public static get observedAttributes(): string[] {
    return ['data-filter', 'data-show-count', 'data-theme-variant', 'data-is-multi'];
  }

  public connectedCallback(): void {
    this._render();
  }

  public attributeChangedCallback(): void {
    if (this.isConnected) { this._render(); }
  }

  private _render(): void {
    const props = this.resolveAttributes();
    const filter: IDataFilterInternal = props.filter;
    const filterName: string = props.filterName || (filter ? filter.filterName : '');
    const instanceId: string = props.instanceId;
    const showCount: boolean = props.showCount === true || props.showCount === 'true';
    const theme: ThemeVariant = props.themeVariant || {};

    const primary: string = (theme.palette && theme.palette.themePrimary) || '#0078d4';
    const text: string = (theme.semanticColors && theme.semanticColors.bodyText) || '#323130';
    const subtle: string = (theme.semanticColors && theme.semanticColors.bodySubtext) || '#605e5c';
    const disabledText: string = (theme.semanticColors && theme.semanticColors.disabledText) || '#a19f9d';

    const groups: IGroupedValue[] = this._group(filter && filter.values ? filter.values : []);

    while (this.firstChild) { this.removeChild(this.firstChild); }

    const list: HTMLDivElement = document.createElement('div');
    list.setAttribute('role', 'group');
    list.setAttribute('aria-label', filter ? filter.displayName : filterName);
    list.style.cssText = 'display:flex;flex-direction:column;gap:2px;padding:4px 0;';

    groups.forEach(g => {
      const row: HTMLLabelElement = document.createElement('label');
      row.title = g.group.extensions.filter(e => !!e).map(e => `.${e}`).join(', ');
      row.style.cssText = `display:flex;align-items:center;gap:8px;padding:4px 8px;border-radius:2px;cursor:${g.disabled ? 'default' : 'pointer'};color:${g.disabled ? disabledText : text};`;

      const checkbox: HTMLInputElement = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = g.selected;
      checkbox.disabled = g.disabled;
      checkbox.style.cssText = `margin:0;width:16px;height:16px;accent-color:${primary};cursor:inherit;`;
      checkbox.addEventListener('change', () => this._dispatch(filterName, instanceId, g, checkbox.checked));

      const icon: HTMLImageElement = document.createElement('img');
      icon.src = getFileTypeIconAsUrl({ extension: g.group.icon, size: 20, imageFileType: 'svg' });
      icon.alt = '';
      icon.width = 20;
      icon.height = 20;
      icon.style.cssText = 'flex:0 0 20px;';

      const label: HTMLSpanElement = document.createElement('span');
      label.textContent = g.group.label;
      label.style.cssText = 'flex:1 1 auto;';

      row.appendChild(checkbox);
      row.appendChild(icon);
      row.appendChild(label);

      if (showCount) {
        const count: HTMLSpanElement = document.createElement('span');
        count.textContent = `(${g.count})`;
        count.style.cssText = `color:${g.disabled ? disabledText : subtle};`;
        row.appendChild(count);
      }

      list.appendChild(row);
    });

    this.appendChild(list);
  }

  /** Groups the values; the order of first occurrence is preserved */
  private _group(values: IDataFilterValueInternal[]): IGroupedValue[] {
    const result: IGroupedValue[] = [];
    const index: { [key: string]: IGroupedValue } = {};

    values.forEach(v => {
      const group: IFileTypeGroup = getFileTypeGroup(v.name || v.value);
      let entry: IGroupedValue = index[group.key];
      if (!entry) {
        entry = { group, members: [], count: 0, selected: false, disabled: true };
        index[group.key] = entry;
        result.push(entry);
      }
      entry.members.push(v);
      entry.count += v.count || 0;
      entry.selected = entry.selected || !!v.selected;
      entry.disabled = entry.disabled && !!v.disabled;
    });

    return result;
  }

  private _dispatch(filterName: string, instanceId: string, g: IGroupedValue, selected: boolean): void {
    const detail: IDataFilterInfo = {
      filterName,
      instanceId,
      operator: FilterConditionOperator.OR,
      filterValues: g.members.map(m => ({
        name: m.name,
        value: m.value,
        selected,
        operator: FilterComparisonOperator.Eq
      }))
    };

    this.dispatchEvent(new CustomEvent(ExtensibilityConstants.EVENT_FILTER_UPDATED, {
      detail,
      bubbles: true,
      cancelable: true
    }));
  }
}
