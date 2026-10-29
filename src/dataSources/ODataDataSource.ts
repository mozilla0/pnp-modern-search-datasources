import { AadHttpClient, HttpClient, HttpClientResponse, IHttpClientOptions, SPHttpClient } from '@microsoft/sp-http';
import { IPropertyPaneField, IPropertyPaneGroup, PropertyPaneDropdown, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import {
  FilterComparisonOperator,
  FilterConditionOperator,
  IDataContext,
  IDataFilter,
  IDataSourceData
} from '@pnp/modern-search-extensibility';
import { applyClientFilters, buildFilterResults } from '../common/ClientDataHelper';
import { ICommonDataSourceProperties, ITokenContext, BaseExtDataSource } from '../common/BaseExtDataSource';
import { firstByPaths, getByPath, JsonObject } from '../common/JsonPath';
import { format } from '../common/Format';
import * as strings from 'ExtDataSourcesStrings';

export type ODataApiType = 'graph' | 'sharepoint' | 'aad' | 'anonymous';

export interface IODataProperties extends ICommonDataSourceProperties {
  apiType: ODataApiType;
  url: string;
  aadResource: string;
  method: 'GET' | 'POST';
  headers: string;
  body: string;
  itemsPath: string;
  totalCountPath: string;
  nextLinkPath: string;
  pagingMode: 'dynamic' | 'static';
  serverSidePaging: boolean;
  maxItems: string;
  filterMode: 'client' | 'odata';
}

interface IPageResult { items: JsonObject[]; total: number; nextLink: string; }

const GRAPH_RESOURCE: string = 'https://graph.microsoft.com';
const MAX_REQUESTS: number = 100;

/**
 * OData data source: Microsoft Graph, SharePoint REST (v1/v2), Entra ID secured APIs
 * and anonymous APIs. Dynamic paging (nextLink or {startRow}/{itemsCountPerPage}) or static paging.
 */
export class ODataDataSource extends BaseExtDataSource<IODataProperties> {

  /** Known page URLs for server-side paging (page → URL) */
  private _pageLinks: { [page: number]: string } = {};
  private _pageLinksSignature: string;
  /** Loaded items for static paging */
  private _staticSignature: string;
  private _staticItems: JsonObject[];

  protected getDefaults(): Partial<IODataProperties> {
    return {
      apiType: 'sharepoint',
      url: '_api/web/lists?$select=Id,Title,Description,ItemCount,LastItemModifiedDate,DefaultViewUrl&$filter=Hidden eq false',
      aadResource: '',
      method: 'GET',
      headers: '{\n  "Accept": "application/json;odata=nometadata"\n}',
      body: '',
      itemsPath: 'value',
      totalCountPath: '@odata.count',
      nextLinkPath: '@odata.nextLink',
      pagingMode: 'static',
      serverSidePaging: false,
      maxItems: '2000',
      filterMode: 'client',
      titleField: 'Title',
      pathField: 'DefaultViewUrl',
      summaryField: 'Description',
      dateField: 'LastItemModifiedDate',
      authorField: '',
      fileTypeField: '',
      idField: 'Id'
    };
  }

  protected resetState(): void {
    this._pageLinks = {};
    this._pageLinksSignature = undefined;
    this._staticSignature = undefined;
    this._staticItems = undefined;
  }

  public async getData(dataContext: IDataContext): Promise<IDataSourceData> {
    if (this.shouldSkipFetch(dataContext)) { return this.emptyResult(); }

    if (this.properties.pagingMode === 'static') {
      return this._getStatic(dataContext);
    }
    return this._getDynamic(dataContext);
  }

  // ---------------------------------------------------------------------------
  // Static paging: load everything, then filter, sort and page in the browser
  // ---------------------------------------------------------------------------

  private async _getStatic(dataContext: IDataContext): Promise<IDataSourceData> {
    // Page, sort and filters are irrelevant for loading
    const ctx: ITokenContext = { ...this.buildTokenContext(dataContext, ''), startRow: 0, pageNumber: 1 };
    const request = await this._buildRequest(ctx, false, '');
    const signature: string = `static|${request.url}|${request.body}|${this.properties.maxItems}`;

    if (this._staticSignature !== signature || !this._staticItems) {
      const cached: JsonObject[] = this.cacheGet<JsonObject[]>(signature);
      this._staticItems = cached || await this._loadAll(request.url, request.body);
      this._staticSignature = signature;
      if (!cached) { this.cacheSet(signature, this._staticItems); }
    }

    return this.processClientSide(this._staticItems, dataContext);
  }

  private async _loadAll(url: string, body: string): Promise<JsonObject[]> {
    const maxItems: number = parseInt(this.properties.maxItems, 10) > 0 ? parseInt(this.properties.maxItems, 10) : 2000;
    const all: JsonObject[] = [];
    let nextUrl: string = url;
    let nextBody: string = body;
    let requests: number = 0;

    while (nextUrl && all.length < maxItems && requests < MAX_REQUESTS) {
      const page: IPageResult = await this._fetchPage(nextUrl, nextBody);
      all.push(...page.items);
      requests++;
      nextUrl = this.properties.serverSidePaging ? page.nextLink : undefined;
      nextBody = undefined;
    }
    return all.slice(0, maxItems);
  }

  // ---------------------------------------------------------------------------
  // Dynamic paging: one request per page
  // ---------------------------------------------------------------------------

  private async _getDynamic(dataContext: IDataContext): Promise<IDataSourceData> {
    const selected: IDataFilter[] = this.getSelectedFilters(dataContext);
    const operator: FilterConditionOperator = this.getFilterOperator(dataContext);
    const serverFilter: boolean = this.properties.filterMode === 'odata';
    const odataFilter: string = serverFilter ? this._buildODataFilter(selected, operator) : '';
    const pageNumber: number = this.getPageNumber(dataContext);
    const rowLimit: number = this.getRowLimit(dataContext);

    let page: IPageResult;

    if (this.properties.serverSidePaging) {
      // Pages are discovered via the nextLink
      const firstCtx: ITokenContext = { ...this.buildTokenContext(dataContext, odataFilter), startRow: 0, pageNumber: 1 };
      const first = await this._buildRequest(firstCtx, true, odataFilter);
      const signature: string = `${first.url}|${first.body}`;

      if (this._pageLinksSignature !== signature) {
        this._pageLinks = { 1: first.url };
        this._pageLinksSignature = signature;
      }

      const targetPage: number = this._pageLinks[pageNumber] ? pageNumber : 1;
      const isFirst: boolean = targetPage === 1;
      const cacheable: boolean = isFirst && selected.length === 0;

      page = cacheable ? this.cacheGet<IPageResult>(signature) : undefined;
      if (!page) {
        page = await this._fetchPage(this._pageLinks[targetPage], isFirst ? first.body : undefined);
        if (cacheable) { this.cacheSet(signature, page); }
      }
      if (page.nextLink) { this._pageLinks[targetPage + 1] = page.nextLink; }

      this.itemCount = page.total > 0
        ? page.total
        : (targetPage - 1) * rowLimit + page.items.length + (page.nextLink ? 1 : 0);

    } else {
      // Pages via {startRow}/{itemsCountPerPage} in URL or body
      const ctx = this.buildTokenContext(dataContext, odataFilter);
      const request = await this._buildRequest(ctx, true, odataFilter);
      const signature: string = `${request.url}|${request.body}`;
      const cacheable: boolean = pageNumber === 1 && selected.length === 0;

      page = cacheable ? this.cacheGet<IPageResult>(signature) : undefined;
      if (!page) {
        page = await this._fetchPage(request.url, request.body);
        if (cacheable) { this.cacheSet(signature, page); }
      }

      if (page.total > 0) {
        this.itemCount = page.total;
      } else if (page.items.length >= rowLimit) {
        this.itemCount = pageNumber * rowLimit + 1; // Total count unknown: offer a next page
      } else {
        this.itemCount = (pageNumber - 1) * rowLimit + page.items.length;
      }
    }

    const items: JsonObject[] = serverFilter ? page.items : applyClientFilters(page.items, selected, operator);
    const configs = dataContext.filters && dataContext.filters.filtersConfiguration ? dataContext.filters.filtersConfiguration : [];

    return {
      items,
      filters: buildFilterResults(page.items, configs, serverFilter ? [] : selected, operator)
    };
  }

  // ---------------------------------------------------------------------------
  // Build and execute the request
  // ---------------------------------------------------------------------------

  private async _buildRequest(
    ctx: ITokenContext,
    applySortAndFilter: boolean,
    odataFilter: string
  ): Promise<{ url: string; body: string }> {

    const rawUrl: string = this.properties.url || '';
    const jsonEscape = (v: string): string => JSON.stringify(v).slice(1, -1);

    let url: string = this.resolveBuiltinTokens(rawUrl, ctx, encodeURIComponent);
    let body: string = this.properties.method === 'POST' ? this.resolveBuiltinTokens(this.properties.body, ctx, jsonEscape) : undefined;

    url = await this.resolvePnPTokens(url);
    if (body) { body = await this.resolvePnPTokens(body); }
    url = this._normalizeUrl(url);

    if (applySortAndFilter) {
      // Append the sort as $orderby unless the URL controls it itself
      if (ctx.sortField && !/\{sortField\}/i.test(rawUrl) && !/[?&]\$orderby=/i.test(url)) {
        url = this._appendQuery(url, `$orderby=${encodeURIComponent(`${ctx.sortField} ${ctx.sortDirection}`)}`);
      }
      // Append the OData filter or combine it with an existing $filter
      if (odataFilter && !/\{filter\}/i.test(rawUrl)) {
        const existing: RegExpMatchArray = url.match(/([?&]\$filter=)([^&]*)/i);
        if (existing) {
          const combined: string = `(${decodeURIComponent(existing[2].replace(/\+/g, ' '))}) and (${odataFilter})`;
          url = url.replace(existing[0], `${existing[1]}${encodeURIComponent(combined)}`);
        } else {
          url = this._appendQuery(url, `$filter=${encodeURIComponent(odataFilter)}`);
        }
      }
    }

    if (!url) { throw new Error(strings.ErrorNoUrl); }
    return { url, body };
  }

  private _normalizeUrl(url: string): string {
    const trimmed: string = (url || '').trim();
    if (/^https?:\/\//i.test(trimmed)) { return trimmed; }

    switch (this.properties.apiType) {
      case 'graph': {
        const path: string = trimmed.replace(/^\/+/, '');
        return /^(v1\.0|beta)\//i.test(path) ? `${GRAPH_RESOURCE}/${path}` : `${GRAPH_RESOURCE}/v1.0/${path}`;
      }
      case 'sharepoint': {
        const origin: string = new URL(this.pageContext.web.absoluteUrl).origin;
        return trimmed.charAt(0) === '/' ? `${origin}${trimmed}` : `${this.pageContext.web.absoluteUrl}/${trimmed}`;
      }
      default:
        return trimmed;
    }
  }

  private _appendQuery(url: string, param: string): string {
    return `${url}${url.indexOf('?') === -1 ? '?' : '&'}${param}`;
  }

  private async _fetchPage(url: string, body?: string): Promise<IPageResult> {
    const json: JsonObject = await this._execute(url, body);

    const itemsPath: string = (this.properties.itemsPath || '').trim();
    const rawItems = itemsPath ? getByPath(json, itemsPath) : json;
    const items: JsonObject[] = Array.isArray(rawItems) ? rawItems : (rawItems && typeof rawItems === 'object' ? [rawItems] : []);

    const total: number = Number(firstByPaths(json, [this.properties.totalCountPath, '@odata.count', 'odata.count'].filter(p => !!p)));
    const next = firstByPaths(json, [this.properties.nextLinkPath, '@odata.nextLink', 'odata.nextLink'].filter(p => !!p));

    return {
      items,
      total: isNaN(total) ? 0 : total,
      nextLink: typeof next === 'string' && next ? new URL(next, url).toString() : undefined
    };
  }

  private async _execute(url: string, body?: string): Promise<JsonObject> {
    const headers: { [key: string]: string } = this.parseJsonSetting(this.properties.headers, strings.HttpHeadersLabel) as { [key: string]: string };
    const has = (name: string): boolean => Object.keys(headers).some(h => h.toLowerCase() === name.toLowerCase());
    const isPost: boolean = !!body;

    if (!has('Accept')) { headers.Accept = 'application/json'; }
    if (isPost && !has('Content-Type')) { headers['Content-Type'] = 'application/json'; }

    const options: IHttpClientOptions = { method: isPost ? 'POST' : 'GET', headers, body: isPost ? body : undefined };
    let response: HttpClientResponse;

    switch (this.properties.apiType) {
      case 'graph': {
        const client: AadHttpClient = await this.getAadClient(GRAPH_RESOURCE);
        response = await client.fetch(url, AadHttpClient.configurations.v1, options);
        break;
      }
      case 'aad': {
        const client: AadHttpClient = await this.getAadClient(this.properties.aadResource);
        response = await client.fetch(url, AadHttpClient.configurations.v1, options);
        break;
      }
      case 'anonymous':
        response = await this.httpClient.fetch(url, HttpClient.configurations.v1, options);
        break;
      default:
        response = await this.spHttpClient.fetch(url, SPHttpClient.configurations.v1, options);
    }

    if (!response.ok) {
      const text: string = await response.text();
      throw new Error(format(strings.ErrorODataFailed, `${response.status} ${response.statusText}`, text.substring(0, 500)));
    }
    return response.json();
  }

  // ---------------------------------------------------------------------------
  // OData $filter from the selected filters
  // ---------------------------------------------------------------------------

  private _buildODataFilter(filters: IDataFilter[], operator: FilterConditionOperator): string {
    const parts: string[] = filters.map(filter => {
      const field: string = filter.filterName.replace(/\./g, '/');
      const expressions: string[] = (filter.values || []).map(v => {
        const value: string = this._odataLiteral(v.value);
        switch (v.operator) {
          case FilterComparisonOperator.Neq: return `${field} ne ${value}`;
          case FilterComparisonOperator.Gt: return `${field} gt ${value}`;
          case FilterComparisonOperator.Geq: return `${field} ge ${value}`;
          case FilterComparisonOperator.Lt: return `${field} lt ${value}`;
          case FilterComparisonOperator.Leq: return `${field} le ${value}`;
          case FilterComparisonOperator.Contains: return `contains(${field},${value})`;
          default: return `${field} eq ${value}`;
        }
      });
      if (expressions.length === 0) { return ''; }
      const joiner: string = filter.operator === FilterConditionOperator.AND ? ' and ' : ' or ';
      return expressions.length === 1 ? expressions[0] : `(${expressions.join(joiner)})`;
    }).filter(p => !!p);

    return parts.join(operator === FilterConditionOperator.OR ? ' or ' : ' and ');
  }

  private _odataLiteral(value: string): string {
    if (/^-?\d+(\.\d+)?$/.test(value) || /^(true|false)$/i.test(value)) { return value; }
    if (/^\d{4}-\d{2}-\d{2}T[\d:.]+(Z|[+-]\d{2}:\d{2})?$/.test(value)) { return value; }
    return `'${String(value).replace(/'/g, '\'\'')}'`;
  }

  // ---------------------------------------------------------------------------
  // Property pane
  // ---------------------------------------------------------------------------

  public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {
    const p = this.properties;
    const urlHelp: { [key in ODataApiType]: string } = {
      graph: strings.UrlHelpGraph,
      sharepoint: strings.UrlHelpSharePoint,
      aad: strings.UrlHelpAad,
      anonymous: strings.UrlHelpAnonymous
    };

    const sourceFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneDropdown('dataSourceProperties.apiType', {
        label: strings.ApiTypeLabel,
        options: [
          { key: 'graph', text: strings.ApiTypeGraph },
          { key: 'sharepoint', text: strings.ApiTypeSharePoint },
          { key: 'aad', text: strings.ApiTypeAad },
          { key: 'anonymous', text: strings.ApiTypeAnonymous }
        ]
      }),
      PropertyPaneTextField('dataSourceProperties.url', {
        label: strings.UrlLabel,
        description: `${urlHelp[p.apiType] || ''}. ${strings.UrlTokensHelp}`,
        multiline: true,
        rows: 4
      })
    ];

    if (p.apiType === 'aad') {
      sourceFields.push(PropertyPaneTextField('dataSourceProperties.aadResource', {
        label: strings.AadResourceLabel,
        description: strings.AadResourceDescription
      }));
    }

    sourceFields.push(
      PropertyPaneDropdown('dataSourceProperties.method', {
        label: strings.HttpMethodLabel,
        options: [{ key: 'GET', text: 'GET' }, { key: 'POST', text: 'POST' }]
      }),
      PropertyPaneTextField('dataSourceProperties.headers', {
        label: strings.HttpHeadersLabel,
        multiline: true,
        rows: 4
      })
    );

    if (p.method === 'POST') {
      sourceFields.push(PropertyPaneTextField('dataSourceProperties.body', {
        label: strings.BodyLabel,
        multiline: true,
        rows: 6
      }));
    }

    const pagingFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneTextField('dataSourceProperties.itemsPath', {
        label: strings.ItemsPathLabel,
        description: strings.ItemsPathDescription
      }),
      PropertyPaneDropdown('dataSourceProperties.pagingMode', {
        label: strings.PagingLabel,
        options: [
          { key: 'dynamic', text: strings.PagingDynamic },
          { key: 'static', text: strings.PagingStatic }
        ]
      }),
      PropertyPaneToggle('dataSourceProperties.serverSidePaging', {
        label: p.pagingMode === 'static' ? strings.ServerPagingStaticLabel : strings.ServerPagingDynamicLabel,
        onText: strings.On,
        offText: strings.Off
      })
    ];

    if (p.serverSidePaging) {
      pagingFields.push(PropertyPaneTextField('dataSourceProperties.nextLinkPath', {
        label: strings.NextLinkPathLabel,
        description: strings.NextLinkPathDescription
      }));
    }

    if (p.pagingMode === 'static') {
      pagingFields.push(PropertyPaneTextField('dataSourceProperties.maxItems', {
        label: strings.MaxItemsLabel
      }));
    } else {
      pagingFields.push(
        PropertyPaneTextField('dataSourceProperties.totalCountPath', {
          label: strings.TotalCountPathLabel,
          description: strings.TotalCountPathDescription
        }),
        PropertyPaneDropdown('dataSourceProperties.filterMode', {
          label: strings.FilterModeLabel,
          options: [
            { key: 'client', text: strings.FilterModeClient },
            { key: 'odata', text: strings.FilterModeOData }
          ]
        })
      );
    }

    return [
      { groupName: strings.GroupODataSource, groupFields: sourceFields },
      { groupName: strings.GroupODataResponse, groupFields: pagingFields },
      ...this.getCommonPropertyPaneGroups()
    ];
  }
}
