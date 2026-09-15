/* tslint:disable:no-any -- SharePoint REST payloads are intentionally dynamic at this boundary. */
import { SPHttpClient } from "@microsoft/sp-http";
import { IBreadcrumbItem } from "../models/IBreadcrumbItem";
import { ILocaleContext, LocaleService } from "../localization/LocaleService";

interface ISharePointCollectionResponse {
  value?: any[];
  d?: {
    results?: any[];
  };
}

/** Loads only active navigation records and resolves the current item's parent chain. */
export class BreadcrumbService {
  public constructor(
    private _spHttpClient: SPHttpClient,
    private _webUrl: string,
  ) {}

  public getBreadcrumbForCurrentPage(
    listId: string,
    locale: ILocaleContext,
  ): Promise<IBreadcrumbItem[]> {
    const select: string =
      "Id,Title,TitleEN,TitleAR,Url,ParentId,DisplayOrder,IsActive,OpenInNewTab,Icon";
    const requestUrl: string =
      this._webUrl +
      "/_api/web/lists(guid'" +
      listId +
      "')/items?$select=" +
      select +
      "&$orderby=DisplayOrder asc,Id asc";

    return this._spHttpClient
      .get(requestUrl, SPHttpClient.configurations.v1)
      .then((response) => this._readResponse(response))
      .then((payload: ISharePointCollectionResponse) => {
        const rawItems: any[] =
          payload.value || (payload.d && payload.d.results) || [];
        // Do not filter IsActive in OData. Existing rows can have a null value when
        // the column was added after the row was created; null is intentionally
        // treated as active by _toItem. An OData filter would hide them before
        // that compatibility rule can be evaluated.
        const items: IBreadcrumbItem[] = rawItems
          .map((raw: any) => this._toItem(raw))
          .filter((item: IBreadcrumbItem) => item.isActive);
        return items;
        // return this._resolveCurrentPath(items, locale);
      });
  }

  private _readResponse(response: Response): Promise<any> {
    if (!response.ok) {
      return response.text().then((body: string) => {
        throw new Error(
          "Breadcrumb list request failed (" + response.status + "): " + body,
        );
      });
    }

    return response.json();
  }

  private _toItem(raw: any): IBreadcrumbItem {
    const rawUrl: any = raw.Url;
    const url: string =
      typeof rawUrl === "string"
        ? rawUrl
        : (rawUrl && (rawUrl.Url || rawUrl.url)) || "";
    const parentId: number = Number(raw.ParentId);

    return {
      id: Number(raw.Id),
      title: raw.Title || "",
      titleEN: raw.TitleEN || "",
      titleAR: raw.TitleAR || "",
      url: this._safeUrl(url),
      parentId: parentId > 0 ? parentId : undefined,
      displayOrder: Number(raw.DisplayOrder) || 0,
      isActive:
        raw.IsActive === undefined ||
        raw.IsActive === null ||
        raw.IsActive === true ||
        raw.IsActive === 1 ||
        raw.IsActive === "1",
      openInNewTab:
        raw.OpenInNewTab === true ||
        raw.OpenInNewTab === 1 ||
        raw.OpenInNewTab === "1",
      icon: raw.Icon || undefined,
    };
  }

  private _resolveCurrentPath(
    items: IBreadcrumbItem[],
    locale: ILocaleContext,
  ): IBreadcrumbItem[] {
    const currentPath: string = this._normaliseUrl(window.location.href);
    let currentItem: IBreadcrumbItem | undefined;
    let bestMatchLength: number = -1;

    items.forEach((item: IBreadcrumbItem) => {
      const itemPath: string = this._normaliseUrl(item.url);
      if (
        itemPath &&
        this._matchesPage(currentPath, itemPath) &&
        itemPath.length > bestMatchLength
      ) {
        currentItem = item;
        bestMatchLength = itemPath.length;
      }
    });

    if (!currentItem) {
      return [];
    }

    const result: IBreadcrumbItem[] = [];
    const visited: { [id: number]: boolean } = {};
    let node: IBreadcrumbItem | undefined = currentItem;

    while (node && !visited[node.id]) {
      visited[node.id] = true;
      if (LocaleService.getItemLabel(node, locale) && node.url) {
        result.unshift(node);
      }
      node = this._findById(items, node.parentId);
    }

    return result;
  }

  private _findById(
    items: IBreadcrumbItem[],
    id: number | undefined,
  ): IBreadcrumbItem | undefined {
    if (!id) {
      return undefined;
    }

    for (let index: number = 0; index < items.length; index++) {
      if (items[index].id === id) {
        return items[index];
      }
    }

    return undefined;
  }

  private _matchesPage(currentPath: string, itemPath: string): boolean {
    if (currentPath === itemPath) {
      return true;
    }

    return currentPath.indexOf(itemPath + "/") === 0;
  }

  private _safeUrl(value: string): string {
    const trimmed: string = (value || "").replace(/^\s+|\s+$/g, "");
    return /^javascript:/i.test(trimmed) ? "" : trimmed;
  }

  private _normaliseUrl(value: string): string {
    const rawValue: string = (value || "").replace(/^\s+|\s+$/g, "");
    if (!rawValue) {
      return "";
    }

    const anchor: HTMLAnchorElement = document.createElement("a");
    // Authors commonly enter "SitePages/Page.aspx" in a navigation list. A
    // browser resolves that relative to the current page, which makes a match
    // depend on where the component is rendered. Resolve relative URLs from the
    // current web instead. Absolute and server-relative URLs keep their meaning.
    if (
      /^(?:[a-z][a-z0-9+.-]*:)?\/\//i.test(rawValue) ||
      rawValue.charAt(0) === "/"
    ) {
      anchor.href = rawValue;
    } else {
      anchor.href =
        this._webUrl.replace(/\/+$/, "") +
        "/" +
        rawValue.replace(/^\.?(?:\/|\\)/, "");
    }

    let path: string = anchor.pathname || rawValue;
    path = path.replace(/[?#].*$/, "").replace(/\/+$/, "");
    return (path || "/").toLowerCase();
  }
}
