/* tslint:disable:no-any -- SharePoint REST payloads are intentionally dynamic at this boundary. */
import { SPHttpClient } from '@microsoft/sp-http';
import { IBreadcrumbConfiguration } from '../models/IBreadcrumbConfiguration';
import { INavigationListReference } from '../models/IBreadcrumbItem';

export const DEFAULT_CONFIGURATION_LIST_TITLE: string = 'Breadcrumb Configuration';
export const DEFAULT_NAVIGATION_LIST_TITLE: string = 'Breadcrumb Navigation';

interface ISharePointList {
  Id: string;
  Title: string;
  BaseTemplate?: number;
  Hidden?: boolean;
  ListItemEntityTypeFullName?: string;
}

interface IFieldDefinition {
  internalName: string;
  type: string;
  payload: any;
}

/**
 * Persists site-scoped settings and provisions/validates the two SharePoint lists.
 * All calls are SharePoint REST calls through the SPFx SPHttpClient.
 */
export class ConfigurationService {
  private _cachedConfiguration: IBreadcrumbConfiguration | undefined;
  private _configurationCacheExpiresAt: number = 0;

  public constructor(
    private _spHttpClient: SPHttpClient,
    private _webUrl: string,
    private _configurationListTitle: string = DEFAULT_CONFIGURATION_LIST_TITLE) {
  }

  public getConfiguration(forceRefresh?: boolean): Promise<IBreadcrumbConfiguration | undefined> {
    if (!forceRefresh && this._cachedConfiguration && Date.now() < this._configurationCacheExpiresAt) {
      return Promise.resolve(this._cachedConfiguration);
    }

    return this._getListByTitle(this._configurationListTitle)
      .then((list: ISharePointList | undefined) => {
        if (!list) {
          return undefined;
        }

        return this._getJson(this._listUrl(list.Id) +
          '/items?$select=Id,Title,NavigationListId,NavigationListTitle&$top=1')
          .then((payload: any) => {
            const rows: any[] = this._collection(payload);
            if (!rows.length || !rows[0].NavigationListId) {
              return undefined;
            }

            const configuration: IBreadcrumbConfiguration = {
              id: Number(rows[0].Id),
              navigationListId: rows[0].NavigationListId,
              navigationListTitle: rows[0].NavigationListTitle || ''
            };
            this._cachedConfiguration = configuration;
            this._configurationCacheExpiresAt = Date.now() + 60000;
            return configuration;
          });
      });
  }

  public getNavigationLists(): Promise<INavigationListReference[]> {
    return this._getJson(this._webUrl +
      '/_api/web/lists?$select=Id,Title,BaseTemplate,Hidden&$filter=Hidden eq false')
      .then((payload: any) => this._collection(payload)
        .filter((list: ISharePointList) => list.BaseTemplate === 100 &&
          list.Title.toLowerCase() !== this._configurationListTitle.toLowerCase())
        .map((list: ISharePointList) => ({ id: list.Id, title: list.Title }))
        .sort((left: INavigationListReference, right: INavigationListReference) =>
          left.title.localeCompare(right.title)));
  }

  public resolveNavigationList(identifier: string): Promise<INavigationListReference> {
    const trimmed: string = (identifier || '').replace(/^\s+|\s+$/g, '');
    if (!trimmed) {
      return Promise.reject(new Error('Select a breadcrumb navigation list.'));
    }

    if (/^[{]?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}[}]?$/i.test(trimmed)) {
      return this._getListById(trimmed.replace(/[{}]/g, ''))
        .then((list: ISharePointList) => ({ id: list.Id, title: list.Title }));
    }

    return this.getNavigationLists().then((lists: INavigationListReference[]) => {
      const desired: string = trimmed.toLowerCase();
      for (let index: number = 0; index < lists.length; index++) {
        if (lists[index].title.toLowerCase() === desired) {
          return lists[index];
        }
      }
      throw new Error('The list "' + trimmed + '" was not found in this site.');
    });
  }

  public validateNavigationList(listId: string): Promise<void> {
    return this._getListById(listId)
      .then((list: ISharePointList) => this._getJson(this._listUrl(list.Id) +
        '/fields?$select=InternalName,TypeAsString'))
      .then((payload: any) => {
        const fields: any[] = this._collection(payload);
        const required: IFieldDefinition[] = this._navigationFields();
        const missing: string[] = [];

        required.forEach((requiredField: IFieldDefinition) => {
          let found: boolean = false;
          fields.forEach((field: any) => {
            if (field.InternalName === requiredField.internalName &&
              (!requiredField.type || field.TypeAsString === requiredField.type)) {
              found = true;
            }
          });
          if (!found) {
            missing.push(requiredField.internalName);
          }
        });

        if (missing.length) {
          throw new Error('The selected list is missing required fields: ' + missing.join(', ') + '.');
        }
      });
  }

  public saveConfiguration(source: INavigationListReference): Promise<IBreadcrumbConfiguration> {
    return this.validateNavigationList(source.id)
      .then(() => this._ensureConfigurationList())
      .then((list: ISharePointList) => this._getJson(this._listUrl(list.Id) +
        '/items?$select=Id,Title&$top=1')
        .then((payload: any) => ({ list: list, existing: this._collection(payload)[0] })))
      .then((result: { list: ISharePointList; existing: any }) => {
        if (!result.list.ListItemEntityTypeFullName) {
          throw new Error('SharePoint did not return the configuration list item entity type.');
        }
        const body: any = {
          __metadata: { type: result.list.ListItemEntityTypeFullName },
          Title: 'DynamicBreadcrumb',
          NavigationListId: source.id,
          NavigationListTitle: source.title
        };
        const collectionUrl: string = this._listUrl(result.list.Id) + '/items';

        if (result.existing) {
          return this._post(collectionUrl + '(' + result.existing.Id + ')', body, 'MERGE')
            .then(() => Number(result.existing.Id));
        }

        return this._post(collectionUrl, body).then(() => 0);
      })
      .then((configurationId: number) => {
        const configuration: IBreadcrumbConfiguration = {
          id: configurationId,
          navigationListId: source.id,
          navigationListTitle: source.title
        };
        this._cachedConfiguration = configuration;
        this._configurationCacheExpiresAt = Date.now() + 60000;
        return configuration;
      });
  }

  public provisionNavigationList(title?: string): Promise<INavigationListReference> {
    const navigationTitle: string = (title || DEFAULT_NAVIGATION_LIST_TITLE).replace(/^\s+|\s+$/g, '') ||
      DEFAULT_NAVIGATION_LIST_TITLE;

    return this._getListByTitle(navigationTitle)
      .then((existing: ISharePointList | undefined) => {
        if (existing) {
          return existing;
        }
        return this._createList(navigationTitle, 'Configuration-driven navigation source for Dynamic Breadcrumb.');
      })
      .then((list: ISharePointList) => this._ensureFields(list, this._navigationFields())
        .then(() => this.validateNavigationList(list.Id))
        .then(() => ({ id: list.Id, title: list.Title })));
  }

  private _ensureConfigurationList(): Promise<ISharePointList> {
    return this._getListByTitle(this._configurationListTitle)
      .then((existing: ISharePointList | undefined) => existing ||
        this._createList(this._configurationListTitle, 'Site-scoped settings for Dynamic Breadcrumb.'))
      .then((list: ISharePointList) => this._ensureFields(list, this._configurationFields())
        .then(() => list));
  }

  private _createList(title: string, description: string): Promise<ISharePointList> {
    return this._post(this._webUrl + '/_api/web/lists', {
      __metadata: { type: 'SP.List' },
      BaseTemplate: 100,
      Title: title,
      Description: description
    }).then(() => this._getListByTitle(title))
      .then((list: ISharePointList | undefined) => {
        if (!list) {
          throw new Error('SharePoint did not return the newly created list "' + title + '".');
        }
        return list;
      });
  }

  private _ensureFields(list: ISharePointList, definitions: IFieldDefinition[]): Promise<void> {
    return this._getJson(this._listUrl(list.Id) + '/fields?$select=InternalName')
      .then((payload: any) => {
        const fields: any[] = this._collection(payload);
        let operation: Promise<void> = Promise.resolve();

        definitions.forEach((definition: IFieldDefinition) => {
          let exists: boolean = false;
          fields.forEach((field: any) => {
            if (field.InternalName === definition.internalName) {
              exists = true;
            }
          });
          if (!exists) {
            operation = operation.then(() => this._post(this._listUrl(list.Id) + '/fields', definition.payload)
              .catch((error: Error) => {
                throw new Error('Unable to create the "' + definition.internalName + '" field: ' + error.message);
              }));
          }
        });

        return operation;
      });
  }

  private _configurationFields(): IFieldDefinition[] {
    return [
      this._textField('NavigationListId'),
      this._textField('NavigationListTitle')
    ];
  }

  private _navigationFields(): IFieldDefinition[] {
    return [
      this._textField('TitleEN'),
      this._textField('TitleAR'),
      this._field('Url', 'URL', 11),
      this._field('ParentId', 'Number', 9, 0),
      this._field('DisplayOrder', 'Number', 9, '0'),
      this._field('IsActive', 'Boolean', 8, '1'),
      this._field('OpenInNewTab', 'Boolean', 8, '0'),
      this._textField('Icon')
    ];
  }

  private _textField(name: string): IFieldDefinition {
    return this._field(name, 'Text', 2);
  }

  /** Uses SharePoint's documented generic field endpoint and FieldTypeKind values. */
  private _field(
    name: string,
    type: string,
    fieldTypeKind: number,
    defaultValue?: string | number): IFieldDefinition {
    const payload: any = {
      __metadata: { type: 'SP.Field' },
      Title: name,
      StaticName: name,
      FieldTypeKind: fieldTypeKind
    };

    if (defaultValue !== undefined) {
      payload.DefaultValue = String(defaultValue);
    }
    return {
      internalName: name,
      type: type,
      payload: payload
    };
  }

  private _getListByTitle(title: string): Promise<ISharePointList | undefined> {
    const escapedTitle: string = title.replace(/'/g, "''");
    const url: string = this._webUrl + "/_api/web/lists/getbytitle('" + escapedTitle +
      "')?$select=Id,Title,BaseTemplate,Hidden,ListItemEntityTypeFullName";

    return this._spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: Response) => {
        if (response.status === 404) {
          return undefined;
        }
        return this._readResponse(response).then((payload: any) => this._single(payload) as ISharePointList);
      });
  }

  private _getListById(id: string): Promise<ISharePointList> {
    return this._getJson(this._listUrl(id) + '?$select=Id,Title,BaseTemplate,Hidden,ListItemEntityTypeFullName')
      .then((payload: any) => this._single(payload) as ISharePointList);
  }

  private _listUrl(id: string): string {
    return this._webUrl + "/_api/web/lists(guid'" + id + "')";
  }

  private _getJson(url: string): Promise<any> {
    return this._spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: Response) => this._readResponse(response));
  }

  private _post(url: string, body: any, method?: string): Promise<void> {
    const headers: Headers = new Headers();
    headers.append('Accept', 'application/json;odata=verbose');
    headers.append('Content-Type', 'application/json;odata=verbose;charset=utf-8');
    // SPHttpClient v1 defaults to OData v4. SharePoint SE list/field provisioning
    // expects the v3 verbose payload shape used below (__metadata: { type: ... }).
    headers.append('OData-Version', '3.0');
    if (method) {
      headers.append('IF-MATCH', '*');
      headers.append('X-HTTP-Method', method);
    }

    return this._spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: headers,
      body: JSON.stringify(body)
    }).then((response: Response) => {
      if (!response.ok) {
        return response.text().then((responseBody: string) => {
          throw new Error('SharePoint update failed (' + response.status + '): ' + responseBody);
        });
      }
    });
  }

  private _readResponse(response: Response): Promise<any> {
    if (!response.ok) {
      return response.text().then((body: string) => {
        throw new Error('SharePoint request failed (' + response.status + '): ' + body);
      });
    }
    return response.json();
  }

  private _collection(payload: any): any[] {
    return payload.value || (payload.d && payload.d.results) || [];
  }

  private _single(payload: any): any {
    return payload.d || payload;
  }
}
