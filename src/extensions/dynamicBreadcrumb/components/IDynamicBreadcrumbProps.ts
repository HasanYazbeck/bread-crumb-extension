import { IBreadcrumbConfiguration } from '../models/IBreadcrumbConfiguration';
import { IBreadcrumbItem } from '../models/IBreadcrumbItem';
import { ILocaleContext } from '../localization/LocaleService';
import { ConfigurationService } from '../services/ConfigurationService';

export interface IDynamicBreadcrumbStrings {
  BreadcrumbAriaLabel: string;
  ConfigureBreadcrumb: string;
  MoreBreadcrumbItems: string;
  ConfigurationRequired: string;
  ConfigurationError: string;
  ConfigurationDialogTitle: string;
  ConfigurationDialogDescription: string;
  NavigationListLabel: string;
  NavigationListPlaceholder: string;
  NavigationListHelp: string;
  CreateNavigationList: string;
  NavigationListNameLabel: string;
  Save: string;
  Cancel: string;
  Saving: string;
  Creating: string;
  ConfigurationSaved: string;
  SelectNavigationList: string;
  RequiredFieldsHelp: string;
}

export interface IDynamicBreadcrumbProps {
  items: IBreadcrumbItem[];
  locale: ILocaleContext;
  strings: IDynamicBreadcrumbStrings;
  canConfigure: boolean;
  configuration: IBreadcrumbConfiguration | undefined;
  configurationService: ConfigurationService;
  configurationOpen: boolean;
  statusMessage?: string;
  onOpenConfiguration: () => void;
  onDismissConfiguration: () => void;
  onConfigurationSaved: () => void;
}
