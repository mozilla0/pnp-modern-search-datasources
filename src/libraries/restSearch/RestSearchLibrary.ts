import { ServiceKey } from '@microsoft/sp-core-library';
import {
  FilterType,
  IAdaptiveCardAction,
  IComponentDefinition,
  IDataSource,
  IDataSourceDefinition,
  IExtensibilityLibrary,
  IFilterControlDefinition,
  ILayoutDefinition,
  ISuggestionProviderDefinition
} from '@pnp/modern-search-extensibility';
import { RestDocumentSearchDataSource } from '../../dataSources/RestDocumentSearchDataSource';
import { ODataDataSource } from '../../dataSources/ODataDataSource';
import { CamlDataSource } from '../../dataSources/CamlDataSource';
import { StaticDataSource } from '../../dataSources/StaticDataSource';
import { FILE_TYPE_FILTER_COMPONENT, FileTypeFilterWebComponent } from '../../components/FileTypeFilterWebComponent';
import * as strings from 'ExtDataSourcesStrings';

/**
 * Extensibility library for PnP Modern Search v4.
 * Important: this file must export this single class only –
 * PnP instantiates the first export of the module.
 */
export class RestSearchLibrary implements IExtensibilityLibrary {

  public name(): string {
    return strings.LibraryName;
  }

  public getCustomDataSources(): IDataSourceDefinition[] {
    return [
      {
        name: strings.DataSourceRestSearch,
        iconName: 'Search',
        key: 'ExtRestDocumentSearch',
        serviceKey: ServiceKey.create<IDataSource>('PnPExtDataSources:RestDocumentSearch', RestDocumentSearchDataSource)
      },
      {
        name: strings.DataSourceOData,
        iconName: 'PlugConnected',
        key: 'ExtOData',
        serviceKey: ServiceKey.create<IDataSource>('PnPExtDataSources:OData', ODataDataSource)
      },
      {
        name: strings.DataSourceCaml,
        iconName: 'List',
        key: 'ExtCaml',
        serviceKey: ServiceKey.create<IDataSource>('PnPExtDataSources:Caml', CamlDataSource)
      },
      {
        name: strings.DataSourceStatic,
        iconName: 'Code',
        key: 'ExtStatic',
        serviceKey: ServiceKey.create<IDataSource>('PnPExtDataSources:Static', StaticDataSource)
      }
    ];
  }

  public getCustomLayouts(): ILayoutDefinition[] {
    return [];
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] { // eslint-disable-line @typescript-eslint/no-explicit-any
    return [
      {
        componentName: FILE_TYPE_FILTER_COMPONENT,
        componentClass: FileTypeFilterWebComponent
      }
    ];
  }

  public getCustomFilterControls(): IFilterControlDefinition[] {
    return [
      {
        key: 'ExtFileTypeFilter',
        name: strings.FilterControlFileType,
        componentName: FILE_TYPE_FILTER_COMPONENT,
        filterType: FilterType.Refiner,
        showOperator: false,
        showApplyButtons: true,
        supportsMultiValues: true,
        supportsValuesCount: true,
        supportsMaxBuckets: true
      }
    ];
  }

  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [];
  }

  public invokeCardAction(action: IAdaptiveCardAction): void {
    // no Adaptive Card actions
  }
}
