import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { IDataContext, IDataSourceData } from '@pnp/modern-search-extensibility';
import { ICommonDataSourceProperties, BaseExtDataSource } from '../common/BaseExtDataSource';
import { getByPath, JsonObject } from '../common/JsonPath';
import * as strings from 'ExtDataSourcesStrings';

export interface IStaticProperties extends ICommonDataSourceProperties {
  staticData: string;
}

function defaultData(): string {
  return JSON.stringify({
    items: [
      { linkName: 'Microsoft Learn', linkHref: 'https://learn.microsoft.com', category: strings.StaticSampleCategoryDocs },
      { linkName: 'Microsoft 365 Admin Center', linkHref: 'https://admin.microsoft.com', category: strings.StaticSampleCategoryAdmin }
    ],
    numberOfLinks: 2
  }, null, 2);
}

/**
 * Static data: items from a JSON object { "items": [...], ... }.
 * Additional properties next to "items" are available in templates.
 */
export class StaticDataSource extends BaseExtDataSource<IStaticProperties> {

  protected getDefaults(): Partial<IStaticProperties> {
    return {
      staticData: defaultData(),
      titleField: 'linkName',
      pathField: 'linkHref',
      summaryField: 'category',
      dateField: '',
      authorField: '',
      fileTypeField: '',
      idField: ''
    };
  }

  public async getData(dataContext: IDataContext): Promise<IDataSourceData> {
    if (this.shouldSkipFetch(dataContext)) { return this.emptyResult(); }

    let json: JsonObject;
    try {
      json = JSON.parse(this.properties.staticData || '{"items":[]}');
    } catch {
      throw new Error(strings.ErrorStaticInvalidJson);
    }

    const items = getByPath(json, 'items');
    if (!Array.isArray(items)) {
      throw new Error(strings.ErrorStaticNoItems);
    }

    const extra: JsonObject = { ...json };
    delete extra.items;
    return this.processClientSide(items, dataContext, extra);
  }

  public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {
    return [
      {
        groupName: strings.GroupStatic,
        groupFields: [
          PropertyPaneTextField('dataSourceProperties.staticData', {
            label: strings.StaticDataLabel,
            description: strings.StaticDataDescription,
            multiline: true,
            rows: 18
          })
        ]
      },
      ...this.getCommonPropertyPaneGroups()
    ];
  }
}
