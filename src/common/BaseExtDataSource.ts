import { ServiceScope } from '@microsoft/sp-core-library';
import { AadHttpClient, AadHttpClientFactory, HttpClient, SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import {
  IPropertyPaneField,
  IPropertyPaneGroup,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import {
  BaseDataSource,
  BuiltinTemplateSlots,
  FilterBehavior,
  FilterConditionOperator,
  IDataContext,
  IDataFilter,
  IDataSourceData,
  ITemplateSlot,
  ITokenService,
  PagingBehavior,
  SortFieldDirection
} from '@pnp/modern-search-extensibility';
import { applyClientFilters, applyClientSort, buildFilterResults, ISortSpec } from './ClientDataHelper';
import { JsonObject } from './JsonPath';
import { buildSortList, ISortFieldConfiguration } from './SortConfiguration';
import { format } from './Format';
import * as strings from 'ExtDataSourcesStrings';

export interface ICommonDataSourceProperties {
  /** Default sort, e.g. "Modified:desc,Title:asc" */
  defaultSort: string;
  /** Fields users may sort by in the sort control */
  userSortFields: string;
  /** Derived from defaultSort/userSortFields; read by the PnP layouts' sort control */
  sortList: ISortFieldConfiguration[];
  cacheEnabled: boolean;
  cacheMinutes: string;
  noFetchOnLoad: boolean;
  titleField: string;
  pathField: string;
  summaryField: string;
  dateField: string;
  authorField: string;
  fileTypeField: string;
  idField: string;
}

export interface ITokenContext {
  terms: string;
  startRow: number;
  rowLimit: number;
  pageNumber: number;
  filter: string;
  sortField: string;
  sortDirection: string;
}

interface ICacheEntry<V> { signature: string; expires: number; value: V; }

export const COMMON_DEFAULTS: Partial<ICommonDataSourceProperties> = {
  defaultSort: '',
  userSortFields: '',
  cacheEnabled: false,
  cacheMinutes: '5',
  noFetchOnLoad: false
};

/**
 * Common base for the additional data sources: services, tokens, sorting, caching,
 * template slots and client-side processing.
 */
export abstract class BaseExtDataSource<T extends ICommonDataSourceProperties> extends BaseDataSource<T> {

  protected pageContext: PageContext;
  protected spHttpClient: SPHttpClient;
  protected httpClient: HttpClient;
  protected itemCount: number = 0;

  private _aadFactory: AadHttpClientFactory;
  private _aadClients: { [resource: string]: AadHttpClient } = {};
  private _tokenService: ITokenService;

  public constructor(serviceScope: ServiceScope) {
    super(serviceScope);
    serviceScope.whenFinished(() => {
      this.pageContext = serviceScope.consume(PageContext.serviceKey);
      this.spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
      this.httpClient = serviceScope.consume(HttpClient.serviceKey);
      this._aadFactory = serviceScope.consume(AadHttpClientFactory.serviceKey);
    });
  }

  /** Default values of the data source */
  protected abstract getDefaults(): Partial<T>;

  /** Called on every property change to reset internal state */
  protected resetState(): void {
    // overridden by subclasses
  }

  public async onInit(): Promise<void> {
    const props = this.properties as unknown as JsonObject;
    const defaults: JsonObject = { ...COMMON_DEFAULTS, ...this.getDefaults() };
    Object.keys(defaults).forEach(key => {
      if (props[key] === undefined || props[key] === null) { props[key] = defaults[key]; }
    });
    this.syncSortList();
  }

  public onPropertyUpdate(): void {
    this.syncSortList();
    this.resetState();
  }

  /** Keeps `sortList` (array expected by the PnP layouts) in sync with the text settings */
  protected syncSortList(): void {
    this.properties.sortList = buildSortList(this.properties.defaultSort, this.properties.userSortFields);
  }

  public getItemCount(): number {
    return this.itemCount;
  }

  public getPagingBehavior(): PagingBehavior {
    return PagingBehavior.Dynamic;
  }

  public getFilterBehavior(): FilterBehavior {
    return FilterBehavior.Dynamic;
  }

  public getSortableFields(): string[] {
    return this.splitList(this.properties.userSortFields);
  }

  public getTemplateSlots(): ITemplateSlot[] {
    const p = this.properties;
    const slots: ITemplateSlot[] = [];
    const add = (slotName: string, field: string): void => { if (field) { slots.push({ slotName, slotField: field }); } };
    add(BuiltinTemplateSlots.Title, p.titleField);
    add(BuiltinTemplateSlots.Path, p.pathField);
    add(BuiltinTemplateSlots.Summary, p.summaryField);
    add(BuiltinTemplateSlots.Date, p.dateField);
    add(BuiltinTemplateSlots.Author, p.authorField);
    add(BuiltinTemplateSlots.FileType, p.fileTypeField);
    add(BuiltinTemplateSlots.Id, p.idField);
    return slots;
  }

  // ---------------------------------------------------------------------------
  // Property pane
  // ---------------------------------------------------------------------------

  protected getCommonPropertyPaneGroups(): IPropertyPaneGroup[] {
    const fields: IPropertyPaneField<unknown>[] = [
      PropertyPaneTextField('dataSourceProperties.defaultSort', {
        label: strings.DefaultSortLabel,
        description: strings.DefaultSortDescription
      }),
      PropertyPaneTextField('dataSourceProperties.userSortFields', {
        label: strings.UserSortFieldsLabel,
        description: strings.UserSortFieldsDescription
      }),
      PropertyPaneToggle('dataSourceProperties.noFetchOnLoad', {
        label: strings.NoFetchOnLoadLabel,
        onText: strings.On,
        offText: strings.Off
      }),
      PropertyPaneToggle('dataSourceProperties.cacheEnabled', {
        label: strings.CacheEnabledLabel,
        onText: strings.On,
        offText: strings.Off
      })
    ];

    if (this.properties.cacheEnabled) {
      fields.push(
        PropertyPaneTextField('dataSourceProperties.cacheMinutes', { label: strings.CacheMinutesLabel }),
        PropertyPaneButton('dataSourceProperties.clearCache', {
          text: strings.ClearCache,
          buttonType: PropertyPaneButtonType.Normal,
          onClick: () => { this.clearCache(); return ''; }
        })
      );
    }

    return [
      { groupName: strings.GroupCommon, groupFields: fields },
      {
        groupName: strings.GroupSlots,
        groupFields: [
          PropertyPaneTextField('dataSourceProperties.titleField', { label: strings.SlotTitle }),
          PropertyPaneTextField('dataSourceProperties.pathField', { label: strings.SlotPath }),
          PropertyPaneTextField('dataSourceProperties.summaryField', { label: strings.SlotSummary }),
          PropertyPaneTextField('dataSourceProperties.dateField', { label: strings.SlotDate }),
          PropertyPaneTextField('dataSourceProperties.authorField', { label: strings.SlotAuthor }),
          PropertyPaneTextField('dataSourceProperties.fileTypeField', { label: strings.SlotFileType }),
          PropertyPaneTextField('dataSourceProperties.idField', { label: strings.SlotId })
        ]
      }
    ];
  }

  // ---------------------------------------------------------------------------
  // Context helpers
  // ---------------------------------------------------------------------------

  protected getSelectedFilters(dataContext: IDataContext): IDataFilter[] {
    return dataContext.filters && dataContext.filters.selectedFilters ? dataContext.filters.selectedFilters : [];
  }

  protected getFilterOperator(dataContext: IDataContext): FilterConditionOperator {
    return dataContext.filters && dataContext.filters.filterOperator ? dataContext.filters.filterOperator : FilterConditionOperator.AND;
  }

  protected getRowLimit(dataContext: IDataContext): number {
    return dataContext.itemsCountPerPage && dataContext.itemsCountPerPage > 0 ? dataContext.itemsCountPerPage : 10;
  }

  protected getPageNumber(dataContext: IDataContext): number {
    return dataContext.pageNumber && dataContext.pageNumber > 0 ? dataContext.pageNumber : 1;
  }

  /** User sort (sort control) or default sort */
  protected getSort(dataContext: IDataContext): ISortSpec[] {
    const sorting = dataContext.sorting;
    if (sorting && sorting.selectedSortFieldName) {
      return [{ field: sorting.selectedSortFieldName, descending: sorting.selectedSortDirection === SortFieldDirection.Descending }];
    }
    return this.splitList(this.properties.defaultSort).map(entry => {
      const parts: string[] = entry.split(':');
      return { field: parts[0].trim(), descending: (parts[1] || '').trim().toLowerCase() === 'desc' };
    }).filter(s => !!s.field);
  }

  /** "Do not fetch on page load": query only once a search term or filter is set */
  protected shouldSkipFetch(dataContext: IDataContext): boolean {
    return !!this.properties.noFetchOnLoad
      && !(dataContext.inputQueryText || '').trim()
      && this.getSelectedFilters(dataContext).length === 0;
  }

  protected emptyResult(): IDataSourceData {
    this.itemCount = 0;
    return { items: [], filters: [] };
  }

  /** Applies filters, counts, sorting and paging in the browser to a loaded set of items */
  protected processClientSide(allItems: JsonObject[], dataContext: IDataContext, extra?: JsonObject): IDataSourceData {
    const selected: IDataFilter[] = this.getSelectedFilters(dataContext);
    const operator: FilterConditionOperator = this.getFilterOperator(dataContext);
    const rowLimit: number = this.getRowLimit(dataContext);
    const start: number = (this.getPageNumber(dataContext) - 1) * rowLimit;

    const filtered: JsonObject[] = applyClientSort(applyClientFilters(allItems, selected, operator), this.getSort(dataContext));
    this.itemCount = filtered.length;

    const configs = dataContext.filters && dataContext.filters.filtersConfiguration ? dataContext.filters.filtersConfiguration : [];
    return {
      ...(extra || {}),
      items: filtered.slice(start, start + rowLimit),
      filters: buildFilterResults(allItems, configs, selected, operator)
    };
  }

  // ---------------------------------------------------------------------------
  // Tokens
  // ---------------------------------------------------------------------------

  protected resolveBuiltinTokens(input: string, ctx: ITokenContext, encode: (v: string) => string): string {
    return (input || '')
      .replace(/\{startRow\}|\{skip\}/gi, String(ctx.startRow))
      .replace(/\{itemsCountPerPage\}|\{top\}|\{rowLimit\}/gi, String(ctx.rowLimit))
      .replace(/\{pageNumber\}/gi, String(ctx.pageNumber))
      .replace(/\{sortField\}/gi, encode(ctx.sortField))
      .replace(/\{sortDirection\}/gi, ctx.sortDirection)
      .replace(/\{filter\}/gi, encode(ctx.filter))
      .replace(/\{searchTermsOrEmpty\}/gi, encode(ctx.terms))
      .replace(/\{searchTerms\}/gi, encode(ctx.terms || '*'))
      .replace(/\{SiteCollection\}/gi, this.pageContext.site.absoluteUrl)
      .replace(/\{Site\}/gi, this.pageContext.web.absoluteUrl);
  }

  /** Resolves PnP tokens such as {User.Email}, {QueryString.x}, {Today}, {Page.<field>} */
  protected async resolvePnPTokens(input: string): Promise<string> {
    if (!input) { return input; }
    if (!this._tokenService && this.serviceKeys && this.serviceKeys.TokenService) {
      try {
        this._tokenService = this.serviceScope.consume(this.serviceKeys.TokenService);
      } catch {
        this._tokenService = undefined;
      }
    }
    return this._tokenService ? this._tokenService.resolveTokens(input) : input;
  }

  protected buildTokenContext(dataContext: IDataContext, filter: string): ITokenContext {
    const rowLimit: number = this.getRowLimit(dataContext);
    const pageNumber: number = this.getPageNumber(dataContext);
    const sort: ISortSpec[] = this.getSort(dataContext);
    return {
      terms: (dataContext.inputQueryText || '').trim(),
      startRow: (pageNumber - 1) * rowLimit,
      rowLimit,
      pageNumber,
      filter,
      sortField: sort.length > 0 ? sort[0].field : '',
      sortDirection: sort.length > 0 && sort[0].descending ? 'desc' : 'asc'
    };
  }

  // ---------------------------------------------------------------------------
  // HTTP
  // ---------------------------------------------------------------------------

  protected async getAadClient(resource: string): Promise<AadHttpClient> {
    const key: string = (resource || '').trim();
    if (!key) { throw new Error(strings.ErrorAadResourceMissing); }
    if (!this._aadClients[key]) {
      this._aadClients[key] = await this._aadFactory.getClient(key);
    }
    return this._aadClients[key];
  }

  protected parseJsonSetting(raw: string, label: string): JsonObject {
    const text: string = (raw || '').trim();
    if (!text) { return {}; }
    try {
      const parsed = JSON.parse(text);
      return typeof parsed === 'object' && parsed !== null ? parsed : {};
    } catch {
      throw new Error(format(strings.InvalidJson, label));
    }
  }

  protected splitList(value: string): string[] {
    return (value || '').split(',').map(s => s.trim()).filter(s => !!s);
  }

  // ---------------------------------------------------------------------------
  // Cache (localStorage, per web part instance)
  // ---------------------------------------------------------------------------

  private get _cacheStorageKey(): string {
    const instanceId: string = this.context && (this.context as JsonObject).instanceId ? (this.context as JsonObject).instanceId : 'default';
    return `pnpExtDataSource_${instanceId}`;
  }

  protected cacheGet<V>(signature: string): V | undefined {
    if (!this.properties.cacheEnabled) { return undefined; }
    try {
      const raw: string = window.localStorage.getItem(this._cacheStorageKey);
      if (!raw) { return undefined; }
      const entry: ICacheEntry<V> = JSON.parse(raw);
      if (entry.signature !== signature || entry.expires < Date.now()) { return undefined; }
      return entry.value;
    } catch {
      return undefined;
    }
  }

  protected cacheSet<V>(signature: string, value: V): void {
    if (!this.properties.cacheEnabled) { return; }
    const minutes: number = parseInt(this.properties.cacheMinutes, 10) > 0 ? parseInt(this.properties.cacheMinutes, 10) : 5;
    const entry: ICacheEntry<V> = { signature, expires: Date.now() + minutes * 60000, value };
    try {
      window.localStorage.setItem(this._cacheStorageKey, JSON.stringify(entry));
    } catch {
      // Storage full or blocked: continue without cache
    }
  }

  protected clearCache(): void {
    try {
      window.localStorage.removeItem(this._cacheStorageKey);
    } catch {
      // ignore
    }
    this.resetState();
  }
}
