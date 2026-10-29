import { ServiceScope } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { IPropertyPaneGroup, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import {
  BaseDataSource,
  BuiltinTemplateSlots,
  FilterBehavior,
  FilterComparisonOperator,
  FilterConditionOperator,
  IDataContext,
  IDataFilter,
  IDataFilterResult,
  IDataFilterResultValue,
  IDataSourceData,
  ITemplateSlot,
  PagingBehavior,
  SortFieldDirection
} from '@pnp/modern-search-extensibility';
import { format } from '../common/Format';
import * as strings from 'ExtDataSourcesStrings';
import { buildSortList, ISortFieldConfiguration } from '../common/SortConfiguration';

export interface IRestDocumentSearchProperties {
  /** KQL query template, e.g. "{searchTerms} IsDocument:1" */
  queryTemplate: string;
  /** Managed properties, comma-separated */
  selectedProperties: string;
  /** Default sort, e.g. "LastModifiedTime:desc,Title:asc" */
  defaultSort: string;
  /** Fields users may sort by in the sort control */
  userSortFields: string;
  /** Derived from defaultSort/userSortFields; read by the PnP layouts' sort control */
  sortList: ISortFieldConfiguration[];
  /** Optional result source (GUID) */
  sourceId: string;
  trimDuplicates: boolean;
  enableQueryRules: boolean;
}

interface ISearchSort { Property: string; Direction: number; }

interface ISearchRequest {
  Querytext: string;
  StartRow: number;
  RowLimit: number;
  SelectProperties: string[];
  TrimDuplicates: boolean;
  EnableQueryRules: boolean;
  ClientType: string;
  SourceId?: string;
  Refiners?: string;
  RefinementFilters?: string[];
  SortList?: ISearchSort[];
}

interface ISearchCell { Key: string; Value: string; ValueType: string; }
interface ISearchRow { Cells: ISearchCell[]; }
interface IRefinerEntry { RefinementName: string; RefinementToken: string; RefinementValue: string; RefinementCount: string; }
interface IRefiner { Name: string; Entries: IRefinerEntry[]; }

const DEFAULT_QUERY_TEMPLATE: string = '{searchTerms} IsDocument:1';
const DEFAULT_SELECT_PROPERTIES: string = [
  'Title', 'Path', 'Filename', 'FileType', 'FileExtension', 'HitHighlightedSummary', 'Author', 'Created',
  'LastModifiedTime', 'ModifiedBy', 'ServerRedirectedURL', 'ServerRedirectedPreviewURL', 'PictureThumbnailURL',
  'SiteId', 'WebId', 'ListId', 'ListItemID', 'UniqueId', 'IsContainer', 'SPWebUrl', 'SiteTitle', 'ParentLink', 'ContentTypeId'
].join(',');
const DEFAULT_REFINER_BUCKETS: number = 100;
/** Number of hits used to compute refiners (same as the built-in PnP source) */
const REFINER_DEEP_HITS: number = 1000000;
const DATE_REFINER_REGEX: RegExp = /^(RefinableDate\w*|LastModifiedTime|Created|ContentModifiedTime|FirstPublishedDate|StartDate|EndDate|.+OWSDATE)$/i;
/** Maximum rows per page allowed by the Search REST API */
const MAX_ROW_LIMIT: number = 500;

/**
 * PnP Modern Search data source that searches documents via the SharePoint Search REST API
 * (/_api/search/postquery). Refiners are computed by the search index over the whole
 * result set, so the 5,000 item list view threshold does not apply.
 */
export class RestDocumentSearchDataSource extends BaseDataSource<IRestDocumentSearchProperties> {

  private _spHttpClient: SPHttpClient;
  private _pageContext: PageContext;
  private _itemCount: number = 0;

  public constructor(serviceScope: ServiceScope) {
    super(serviceScope);
    serviceScope.whenFinished(() => {
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
    });
  }

  public async onInit(): Promise<void> {
    if (!this.properties.queryTemplate) { this.properties.queryTemplate = DEFAULT_QUERY_TEMPLATE; }
    if (!this.properties.selectedProperties) { this.properties.selectedProperties = DEFAULT_SELECT_PROPERTIES; }
    // Migration: up to 2.3.1 the default sort was stored as text in "sortList"
    const legacySort: unknown = this.properties.sortList;
    if (typeof legacySort === 'string' && this.properties.defaultSort === undefined) {
      this.properties.defaultSort = legacySort;
    }
    if (this.properties.defaultSort === undefined) { this.properties.defaultSort = ''; }
    if (this.properties.userSortFields === undefined) { this.properties.userSortFields = ''; }
    this._syncSortList();
    if (this.properties.sourceId === undefined) { this.properties.sourceId = ''; }
    if (this.properties.trimDuplicates === undefined) { this.properties.trimDuplicates = false; }
    if (this.properties.enableQueryRules === undefined) { this.properties.enableQueryRules = false; }
  }

  public async getData(dataContext: IDataContext): Promise<IDataSourceData> {
    const rowLimit: number = Math.min(dataContext.itemsCountPerPage || 10, MAX_ROW_LIMIT);
    const pageNumber: number = dataContext.pageNumber && dataContext.pageNumber > 0 ? dataContext.pageNumber : 1;

    const request: ISearchRequest = {
      Querytext: this._buildQueryText(dataContext.inputQueryText),
      StartRow: (pageNumber - 1) * rowLimit,
      RowLimit: rowLimit,
      SelectProperties: this._splitList(this.properties.selectedProperties || DEFAULT_SELECT_PROPERTIES),
      TrimDuplicates: !!this.properties.trimDuplicates,
      EnableQueryRules: !!this.properties.enableQueryRules,
      ClientType: 'PnPModernSearch'
    };

    if (this.properties.sourceId && this.properties.sourceId.trim()) {
      request.SourceId = this.properties.sourceId.trim();
    }

    // Refiners from the configuration of the Search Filters web part
    const filters = dataContext.filters;
    if (filters && filters.filtersConfiguration && filters.filtersConfiguration.length > 0) {
      const specs: string[] = [];
      const seen: string[] = [];
      filters.filtersConfiguration.forEach(cfg => {
        if (cfg.filterName && seen.indexOf(cfg.filterName) === -1) {
          seen.push(cfg.filterName);
          specs.push(this._buildRefinerSpec(cfg.filterName, cfg.maxBuckets));
        }
      });
      if (specs.length > 0) { request.Refiners = specs.join(','); }
    }

    // Selected filter values → FQL RefinementFilters
    if (filters && filters.selectedFilters && filters.selectedFilters.length > 0) {
      const refinementFilters: string[] = this._buildRefinementFilters(filters.selectedFilters, filters.filterOperator);
      if (refinementFilters.length > 0) { request.RefinementFilters = refinementFilters; }
    }

    const sortList: ISearchSort[] = this._buildSortList(dataContext);
    if (sortList.length > 0) { request.SortList = sortList; }

    const response: SPHttpClientResponse = await this._spHttpClient.post(
      `${this._pageContext.web.absoluteUrl}/_api/search/postquery`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'odata-version': '3.0'
        },
        body: JSON.stringify({ request })
      }
    );

    if (!response.ok) {
      const errorText: string = await response.text();
      throw new Error(format(strings.ErrorSearchFailed, response.status, errorText));
    }

    const json = await response.json();
    const primary = json && json.PrimaryQueryResult;

    if (!primary || !primary.RelevantResults) {
      this._itemCount = 0;
      return { items: [], filters: [] };
    }

    const relevant = primary.RelevantResults;
    this._itemCount = relevant.TotalRows || 0;

    const rows: ISearchRow[] = relevant.Table && relevant.Table.Rows ? relevant.Table.Rows : [];
    const items = rows.map(row => this._rowToItem(row.Cells));

    const refiners: IRefiner[] = primary.RefinementResults && primary.RefinementResults.Refiners
      ? primary.RefinementResults.Refiners
      : [];
    const filterResults: IDataFilterResult[] = refiners.map(r => this._mapRefiner(r));

    return { items, filters: filterResults };
  }

  public getItemCount(): number {
    return this._itemCount;
  }

  public getPagingBehavior(): PagingBehavior {
    return PagingBehavior.Dynamic;
  }

  public getFilterBehavior(): FilterBehavior {
    return FilterBehavior.Dynamic;
  }

  public getSortableFields(): string[] {
    return this._splitList(this.properties.userSortFields);
  }

  public onPropertyUpdate(): void {
    this._syncSortList();
  }

  /** Keeps `sortList` (array expected by the PnP layouts) in sync with the text settings */
  private _syncSortList(): void {
    this.properties.sortList = buildSortList(this.properties.defaultSort, this.properties.userSortFields);
  }

  public getTemplateSlots(): ITemplateSlot[] {
    return [
      { slotName: BuiltinTemplateSlots.Title, slotField: 'Title' },
      { slotName: BuiltinTemplateSlots.Path, slotField: 'Path' },
      { slotName: BuiltinTemplateSlots.Summary, slotField: 'HitHighlightedSummary' },
      { slotName: BuiltinTemplateSlots.FileType, slotField: 'FileType' },
      { slotName: BuiltinTemplateSlots.PreviewImageUrl, slotField: 'PictureThumbnailURL' },
      { slotName: BuiltinTemplateSlots.PreviewUrl, slotField: 'ServerRedirectedPreviewURL' },
      { slotName: BuiltinTemplateSlots.Author, slotField: 'Author' },
      { slotName: BuiltinTemplateSlots.Date, slotField: 'LastModifiedTime' },
      { slotName: BuiltinTemplateSlots.SiteId, slotField: 'SiteId' },
      { slotName: BuiltinTemplateSlots.WebId, slotField: 'WebId' },
      { slotName: BuiltinTemplateSlots.ListId, slotField: 'ListId' },
      { slotName: BuiltinTemplateSlots.ItemId, slotField: 'ListItemID' },
      { slotName: BuiltinTemplateSlots.IsFolder, slotField: 'IsContainer' }
    ];
  }

  public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {
    return [
      {
        groupName: strings.GroupRestSearch,
        groupFields: [
          PropertyPaneTextField('dataSourceProperties.queryTemplate', {
            label: strings.QueryTemplateLabel,
            description: strings.QueryTemplateDescription,
            multiline: true,
            rows: 3
          }),
          PropertyPaneTextField('dataSourceProperties.selectedProperties', {
            label: strings.SelectedPropertiesLabel,
            multiline: true,
            rows: 5
          }),
          PropertyPaneTextField('dataSourceProperties.defaultSort', {
            label: strings.DefaultSortLabel,
            description: strings.RestSortDescription
          }),
          PropertyPaneTextField('dataSourceProperties.userSortFields', {
            label: strings.UserSortFieldsLabel,
            description: strings.UserSortFieldsDescription
          }),
          PropertyPaneTextField('dataSourceProperties.sourceId', {
            label: strings.SourceIdLabel
          }),
          PropertyPaneToggle('dataSourceProperties.trimDuplicates', {
            label: strings.TrimDuplicatesLabel,
            onText: strings.On,
            offText: strings.Off
          }),
          PropertyPaneToggle('dataSourceProperties.enableQueryRules', {
            label: strings.EnableQueryRulesLabel,
            onText: strings.On,
            offText: strings.Off
          })
        ]
      }
    ];
  }

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------

  private _buildQueryText(inputQueryText: string): string {
    const terms: string = (inputQueryText || '').trim() || '*';
    const query: string = (this.properties.queryTemplate || DEFAULT_QUERY_TEMPLATE)
      .replace(/\{searchTerms\}/gi, terms)
      .replace(/\{Site\}/gi, this._pageContext.web.absoluteUrl)
      .replace(/\{SiteCollection\}/gi, this._pageContext.site.absoluteUrl)
      .trim();
    return query || '*';
  }

  private _buildRefinerSpec(name: string, maxBuckets?: number): string {
    // Date fields: fixed intervals (last year, 3 months, month, week, 24 h)
    if (DATE_REFINER_REGEX.test(name)) {
      const now: Date = new Date();
      const shift = (months: number, days: number, hours: number): string => {
        const d: Date = new Date(now.getTime());
        d.setMonth(d.getMonth() - months);
        d.setDate(d.getDate() - days);
        d.setHours(d.getHours() - hours);
        d.setMinutes(d.getMinutes() - 1);
        return d.toISOString();
      };
      return `${name}(discretize=manual/${shift(12, 0, 0)}/${shift(3, 0, 0)}/${shift(1, 0, 0)}/${shift(0, 7, 0)}/${shift(0, 0, 24)}/${now.toISOString()},deephits=${REFINER_DEEP_HITS})`;
    }
    const max: number = maxBuckets && maxBuckets > 0 ? maxBuckets : DEFAULT_REFINER_BUCKETS;
    // deephits: compute refiners over the whole result set, not only the top hits
    return `${name}(filter=${max}/0/*,sort=frequency/descending,deephits=${REFINER_DEEP_HITS})`;
  }

  private _buildRefinementFilters(selected: IDataFilter[], operator: FilterConditionOperator): string[] {
    const perFilter: string[] = [];

    selected.forEach(filter => {
      const expressions: string[] = (filter.values || [])
        .map(v => this._buildValueExpression(filter.filterName, v.value, v.operator))
        .filter(e => !!e);
      if (expressions.length > 0) {
        perFilter.push(this._combine(expressions, filter.operator));
      }
    });

    if (perFilter.length === 0) { return []; }

    // Multiple RefinementFilters entries are combined with AND by SharePoint
    if (operator === FilterConditionOperator.OR && perFilter.length > 1) {
      return [`or(${perFilter.join(',')})`];
    }
    return perFilter;
  }

  private _combine(expressions: string[], operator: FilterConditionOperator): string {
    if (expressions.length === 1) { return expressions[0]; }
    const op: string = operator === FilterConditionOperator.AND ? 'and' : 'or';
    return `${op}(${expressions.join(',')})`;
  }

  private _buildValueExpression(name: string, value: string, operator: FilterComparisonOperator): string {
    if (!value) { return ''; }

    // Keep refinement tokens (ǂǂ…) and complete range() expressions as they are
    const isToken: boolean = value.indexOf('ǂǂ') === 0 || value.indexOf('range(') === 0;
    if (isToken) {
      return operator === FilterComparisonOperator.Neq ? `not(${name}:${value})` : `${name}:${value}`;
    }

    const fql: string = this._toFqlValue(value);
    switch (operator) {
      case FilterComparisonOperator.Neq: return `not(${name}:${fql})`;
      case FilterComparisonOperator.Gt: return `${name}:range(${fql},max,from="gt")`;
      case FilterComparisonOperator.Geq: return `${name}:range(${fql},max,from="ge")`;
      case FilterComparisonOperator.Lt: return `${name}:range(min,${fql},to="lt")`;
      case FilterComparisonOperator.Leq: return `${name}:range(min,${fql},to="le")`;
      default: return `${name}:${fql}`; // Eq / Contains
    }
  }

  private _toFqlValue(value: string): string {
    if (/^\d{4}-\d{2}-\d{2}/.test(value)) { return `datetime("${value}")`; }
    if (/^-?\d+(\.\d+)?$/.test(value)) { return value; }
    if (value.charAt(0) === '"' && value.charAt(value.length - 1) === '"') { return value; }
    return `"${value.replace(/"/g, '\\"')}"`;
  }

  private _buildSortList(dataContext: IDataContext): ISearchSort[] {
    // Dynamic sort (e.g. from the sort control of the results)
    const sorting = dataContext.sorting;
    if (sorting && sorting.selectedSortFieldName) {
      const ascending: boolean = sorting.selectedSortDirection === SortFieldDirection.Ascending;
      return [{ Property: sorting.selectedSortFieldName, Direction: ascending ? 0 : 1 }];
    }

    // Default sort from the properties
    return this._splitList(this.properties.defaultSort || '').map(entry => {
      const parts: string[] = entry.split(':');
      const direction: string = (parts[1] || 'asc').trim().toLowerCase();
      return { Property: parts[0].trim(), Direction: direction === 'desc' ? 1 : 0 };
    }).filter(s => !!s.Property);
  }

  private _mapRefiner(refiner: IRefiner): IDataFilterResult {
    const values: IDataFilterResultValue[] = (refiner.Entries || []).map(e => ({
      name: e.RefinementName,
      value: e.RefinementToken,
      count: parseInt(e.RefinementCount, 10) || 0
    }));
    return { filterName: refiner.Name, values };
  }

  private _rowToItem(cells: ISearchCell[]): { [key: string]: string } {
    const item: { [key: string]: string } = {};
    (cells || []).forEach(c => { item[c.Key] = c.Value; });
    return item;
  }

  private _splitList(value: string): string[] {
    return (value || '').split(',').map(s => s.trim()).filter(s => !!s);
  }
}
