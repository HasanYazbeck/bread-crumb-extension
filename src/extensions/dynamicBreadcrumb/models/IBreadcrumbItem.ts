/** A normalized navigation record from the configured SharePoint list. */
export interface IBreadcrumbItem {
  id: number;
  title: string;
  titleEN: string;
  titleAR: string;
  url: string;
  parentId?: number;
  displayOrder: number;
  isActive: boolean;
  openInNewTab: boolean;
  icon?: string;
}

/** A list that can be used as the breadcrumb navigation source. */
export interface INavigationListReference {
  id: string;
  title: string;
}
