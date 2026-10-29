import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IPropertyPaneGroup, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import {
  FilterComparisonOperator,
  FilterConditionOperator,
  IDataContext,
  IDataFilter,
  IDataFilterValue,
  IDataSourceData
} from '@pnp/modern-search-extensibility';
import { buildFilterResults, ISortSpec } from '../common/ClientDataHelper';
import { ICommonDataSourceProperties, ITokenContext, BaseExtDataSource } from '../common/BaseExtDataSource';
import { JsonObject } from '../common/JsonPath';
import { format } from '../common/Format';
import * as strings from 'ExtDataSourcesStrings';

export interface ICamlProperties extends ICommonDataSourceProperties {
  /** Site URL (empty = current site), absolute or server-relative */
  siteUrl: string;
  /** List title, list GUID or server-relative list URL */
  list: string;
  viewFields: string;
  searchFields: string;
  recursive: boolean;
  camlQuery: string;
}

interface ICamlPage { items: JsonObject[]; nextHref: string; }

const DEFAULT_CAML: string = `<View>
  <Query>
    <OrderBy>
      <FieldRef Name="Modified" Ascending="FALSE" />
    </OrderBy>
  </Query>
</View>`;

/**
 * SharePoint CAML data source via RenderListDataAsStream. Supports tokens in CAML,
 * search terms over configurable fields, PnP filters as additional CAML conditions,
 * sorting and paging via NextHref.
 */
export class CamlDataSource extends BaseExtDataSource<ICamlProperties> {

  private _pageLinks: { [page: number]: string } = {};
  private _pageLinksSignature: string;

  protected getDefaults(): Partial<ICamlProperties> {
    return {
      siteUrl: '',
      list: strings.DefaultListTitle,
      viewFields: 'ID,Title,FileLeafRef,FileRef,File_x0020_Type,UniqueId,FSObjType,Modified,Created,Editor,Author',
      searchFields: 'Title,FileLeafRef',
      recursive: true,
      camlQuery: DEFAULT_CAML,
      titleField: 'FileLeafRef',
      pathField: 'AbsoluteFileRef',
      summaryField: 'Title',
      dateField: 'Modified',
      authorField: 'Editor[0].title',
      fileTypeField: 'File_x0020_Type',
      idField: 'UniqueId'
    };
  }

  protected resetState(): void {
    this._pageLinks = {};
    this._pageLinksSignature = undefined;
  }

  public async getData(dataContext: IDataContext): Promise<IDataSourceData> {
    if (this.shouldSkipFetch(dataContext)) { return this.emptyResult(); }

    const selected: IDataFilter[] = this.getSelectedFilters(dataContext);
    const operator: FilterConditionOperator = this.getFilterOperator(dataContext);
    const pageNumber: number = this.getPageNumber(dataContext);
    const rowLimit: number = this.getRowLimit(dataContext);

    const viewXml: string = await this._buildViewXml(dataContext, selected, operator, rowLimit);
    const endpoint: string = this._getEndpoint();
    const signature: string = `${endpoint}|${viewXml}`;

    if (this._pageLinksSignature !== signature) {
      this._pageLinks = { 1: '' };
      this._pageLinksSignature = signature;
    }

    const targetPage: number = this._pageLinks[pageNumber] !== undefined ? pageNumber : 1;
    const cacheable: boolean = targetPage === 1 && selected.length === 0;

    let page: ICamlPage = cacheable ? this.cacheGet<ICamlPage>(signature) : undefined;
    if (!page) {
      page = await this._fetch(endpoint, this._pageLinks[targetPage], viewXml);
      if (cacheable) { this.cacheSet(signature, page); }
    }
    if (page.nextHref) { this._pageLinks[targetPage + 1] = page.nextHref; }

    this.itemCount = (targetPage - 1) * rowLimit + page.items.length + (page.nextHref ? 1 : 0);

    const configs = dataContext.filters && dataContext.filters.filtersConfiguration ? dataContext.filters.filtersConfiguration : [];
    return {
      items: page.items,
      filters: buildFilterResults(page.items, configs, [], operator)
    };
  }

  // ---------------------------------------------------------------------------
  // Build CAML
  // ---------------------------------------------------------------------------

  private async _buildViewXml(
    dataContext: IDataContext, selected: IDataFilter[], operator: FilterConditionOperator, rowLimit: number
  ): Promise<string> {

    const ctx: ITokenContext = this.buildTokenContext(dataContext, '');
    let xml: string = this.resolveBuiltinTokens(this.properties.camlQuery || DEFAULT_CAML, ctx, this._xmlEscape);
    xml = (await this.resolvePnPTokens(xml)).trim();

    if (!/^<View[\s>]/i.test(xml)) {
      xml = /^<Query[\s>]/i.test(xml) ? `<View>${xml}</View>` : `<View><Query>${xml}</Query></View>`;
    }

    const doc: Document = new DOMParser().parseFromString(xml, 'text/xml');
    if (doc.getElementsByTagName('parsererror').length > 0) {
      throw new Error(strings.ErrorCamlInvalidXml);
    }

    const view: Element = doc.documentElement;
    if (this.properties.recursive) { view.setAttribute('Scope', 'RecursiveAll'); }

    const query: Element = this._child(view, 'Query') || view.appendChild(doc.createElement('Query'));

    // Conditions: existing Where + search terms + selected filters
    const conditions: string[] = [];
    const where: Element = this._child(query, 'Where');
    if (where && where.firstElementChild) {
      conditions.push(new XMLSerializer().serializeToString(where.firstElementChild));
    }
    if (where) { query.removeChild(where); }

    const searchCondition: string = this._buildSearchCondition(ctx.terms);
    if (searchCondition) { conditions.push(searchCondition); }

    const filterCondition: string = this._buildFilterCondition(selected, operator);
    if (filterCondition) { conditions.push(filterCondition); }

    if (conditions.length > 0) {
      const whereDoc: Document = new DOMParser().parseFromString(`<Where>${this._combine(conditions, 'And')}</Where>`, 'text/xml');
      query.insertBefore(doc.importNode(whereDoc.documentElement, true), query.firstChild);
    }

    // User sort or default sort replaces OrderBy
    const sort: ISortSpec[] = this.getSort(dataContext);
    if (sort.length > 0) {
      const orderBy: Element = this._child(query, 'OrderBy');
      if (orderBy) { query.removeChild(orderBy); }
      const newOrderBy: Element = doc.createElement('OrderBy');
      sort.forEach(s => {
        const ref: Element = doc.createElement('FieldRef');
        ref.setAttribute('Name', s.field);
        ref.setAttribute('Ascending', s.descending ? 'FALSE' : 'TRUE');
        newOrderBy.appendChild(ref);
      });
      query.appendChild(newOrderBy);
    }

    // ViewFields
    const fields: string[] = this.splitList(this.properties.viewFields);
    if (fields.length > 0) {
      const existing: Element = this._child(view, 'ViewFields');
      if (existing) { view.removeChild(existing); }
      const viewFields: Element = doc.createElement('ViewFields');
      fields.forEach(f => {
        const ref: Element = doc.createElement('FieldRef');
        ref.setAttribute('Name', f);
        viewFields.appendChild(ref);
      });
      view.appendChild(viewFields);
    }

    // RowLimit
    const oldLimit: Element = this._child(view, 'RowLimit');
    if (oldLimit) { view.removeChild(oldLimit); }
    const limit: Element = doc.createElement('RowLimit');
    limit.setAttribute('Paged', 'TRUE');
    limit.textContent = String(rowLimit);
    view.appendChild(limit);

    return new XMLSerializer().serializeToString(doc);
  }

  private _buildSearchCondition(terms: string): string {
    const fields: string[] = this.splitList(this.properties.searchFields);
    if (!terms || terms === '*' || fields.length === 0) { return ''; }
    const value: string = this._xmlEscape(terms.replace(/\*/g, ''));
    if (!value) { return ''; }
    return this._combine(
      fields.map(f => `<Contains><FieldRef Name="${this._xmlEscape(f)}" /><Value Type="Text">${value}</Value></Contains>`),
      'Or'
    );
  }

  private _buildFilterCondition(filters: IDataFilter[], operator: FilterConditionOperator): string {
    const perFilter: string[] = filters.map(filter => {
      const field: string = filter.filterName.split(/[.[]/)[0];
      const expressions: string[] = (filter.values || []).map(v => this._valueCondition(field, v));
      return expressions.length > 0
        ? this._combine(expressions, filter.operator === FilterConditionOperator.AND ? 'And' : 'Or')
        : '';
    }).filter(e => !!e);

    return perFilter.length > 0
      ? this._combine(perFilter, operator === FilterConditionOperator.OR ? 'Or' : 'And')
      : '';
  }

  private _valueCondition(field: string, v: IDataFilterValue): string {
    const raw: string = String(v.value);
    let type: string = 'Text';
    let extra: string = '';
    if (/^\d{4}-\d{2}-\d{2}/.test(raw)) {
      type = 'DateTime';
      extra = ' IncludeTimeValue="TRUE"';
    } else if (/^-?\d+(\.\d+)?$/.test(raw)) {
      type = 'Number';
    }

    const op: string = {
      [FilterComparisonOperator.Eq]: 'Eq',
      [FilterComparisonOperator.Neq]: 'Neq',
      [FilterComparisonOperator.Gt]: 'Gt',
      [FilterComparisonOperator.Lt]: 'Lt',
      [FilterComparisonOperator.Geq]: 'Geq',
      [FilterComparisonOperator.Leq]: 'Leq',
      [FilterComparisonOperator.Contains]: 'Contains'
    }[v.operator !== undefined ? v.operator : FilterComparisonOperator.Eq] || 'Eq';

    return `<${op}><FieldRef Name="${this._xmlEscape(field)}" /><Value Type="${type}"${extra}>${this._xmlEscape(raw)}</Value></${op}>`;
  }

  /** CAML only supports binary And/Or: nest the conditions */
  private _combine(conditions: string[], op: 'And' | 'Or'): string {
    if (conditions.length === 0) { return ''; }
    return conditions.reduce((acc, c) => (acc ? `<${op}>${acc}${c}</${op}>` : c), '');
  }

  private _child(parent: Element, name: string): Element {
    for (let i = 0; i < parent.children.length; i++) {
      if (parent.children[i].nodeName.toLowerCase() === name.toLowerCase()) { return parent.children[i]; }
    }
    return undefined;
  }

  private _xmlEscape(value: string): string {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  // ---------------------------------------------------------------------------
  // HTTP
  // ---------------------------------------------------------------------------

  private _getWebUrl(): string {
    const site: string = (this.properties.siteUrl || '').trim().replace(/\/+$/, '');
    if (!site) { return this.pageContext.web.absoluteUrl; }
    if (/^https?:\/\//i.test(site)) { return site; }
    return `${new URL(this.pageContext.web.absoluteUrl).origin}${site.charAt(0) === '/' ? '' : '/'}${site}`;
  }

  private _getEndpoint(): string {
    const list: string = (this.properties.list || '').trim();
    if (!list) { throw new Error(strings.ErrorNoList); }

    const web: string = this._getWebUrl();
    let listPart: string;
    if (/^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/i.test(list)) {
      listPart = `lists(guid'${list.replace(/[{}]/g, '')}')`;
    } else if (list.charAt(0) === '/') {
      listPart = `GetList('${encodeURIComponent(list.replace(/'/g, '\'\''))}')`;
    } else {
      listPart = `lists/GetByTitle('${encodeURIComponent(list.replace(/'/g, '\'\''))}')`;
    }
    return `${web}/_api/web/${listPart}/RenderListDataAsStream`;
  }

  private async _fetch(endpoint: string, nextHref: string, viewXml: string): Promise<ICamlPage> {
    const url: string = nextHref ? `${endpoint}${nextHref.charAt(0) === '?' ? nextHref : `?${nextHref}`}` : endpoint;

    const response: SPHttpClientResponse = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': '3.0'
      },
      body: JSON.stringify({
        parameters: {
          RenderOptions: 2,
          ViewXml: viewXml,
          DatesInUtc: true
        }
      })
    });

    if (!response.ok) {
      const text: string = await response.text();
      throw new Error(format(strings.ErrorCamlFailed, response.status, text.substring(0, 500)));
    }

    const json: JsonObject = await response.json();
    const origin: string = new URL(this._getWebUrl()).origin;
    const items: JsonObject[] = (json.Row || []).map((row: JsonObject) => ({
      ...row,
      AbsoluteFileRef: row.FileRef ? `${origin}${row.FileRef}` : undefined
    }));

    return { items, nextHref: json.NextHref || undefined };
  }

  // ---------------------------------------------------------------------------
  // Property pane
  // ---------------------------------------------------------------------------

  public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {
    return [
      {
        groupName: strings.GroupCamlList,
        groupFields: [
          PropertyPaneTextField('dataSourceProperties.siteUrl', {
            label: strings.SiteUrlLabel,
            description: strings.SiteUrlDescription
          }),
          PropertyPaneTextField('dataSourceProperties.list', {
            label: strings.ListLabel,
            description: strings.ListDescription
          }),
          PropertyPaneTextField('dataSourceProperties.viewFields', {
            label: strings.ViewFieldsLabel,
            multiline: true,
            rows: 3
          }),
          PropertyPaneToggle('dataSourceProperties.recursive', {
            label: strings.RecursiveLabel,
            onText: strings.On,
            offText: strings.Off
          })
        ]
      },
      {
        groupName: strings.GroupCamlQuery,
        groupFields: [
          PropertyPaneTextField('dataSourceProperties.camlQuery', {
            label: strings.CamlLabel,
            description: strings.CamlDescription,
            multiline: true,
            rows: 12
          }),
          PropertyPaneTextField('dataSourceProperties.searchFields', {
            label: strings.SearchFieldsLabel,
            description: strings.SearchFieldsDescription
          })
        ]
      },
      ...this.getCommonPropertyPaneGroups()
    ];
  }
}
