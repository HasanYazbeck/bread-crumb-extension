/* tslint:disable:no-any -- React.createElement must use the React 15-compatible element shape. */
import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import { SPPermission } from "@microsoft/sp-page-context";
import * as React from "react";
import * as ReactDom from "react-dom";

import { DynamicBreadcrumb } from "./components/DynamicBreadcrumb";
import { IDynamicBreadcrumbStrings } from "./components/IDynamicBreadcrumbProps";
import { IBreadcrumbConfiguration } from "./models/IBreadcrumbConfiguration";
import { IBreadcrumbItem } from "./models/IBreadcrumbItem";
import { LocaleService, ILocaleContext } from "./localization/LocaleService";
import { BreadcrumbService } from "./services/BreadcrumbService";
import { ConfigurationService } from "./services/ConfigurationService";
import { ThemeService } from "./theming/ThemeService";

import * as strings from "DynamicBreadcrumbApplicationCustomizerStrings";

const LOG_SOURCE: string = "DynamicBreadcrumb";

/** Properties are optional because the site-scoped configuration list owns runtime settings. */
export interface IDynamicBreadcrumbApplicationCustomizerProperties {
  configurationListTitle?: string;
}

/** Renders a configurable, list-driven breadcrumb into the modern page Top placeholder. */
export default class DynamicBreadcrumbApplicationCustomizer extends BaseApplicationCustomizer<IDynamicBreadcrumbApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent;
  private _configurationService: ConfigurationService;
  private _breadcrumbService: BreadcrumbService;
  private _themeService: ThemeService = new ThemeService();
  private _locale: ILocaleContext;
  private _configuration: IBreadcrumbConfiguration;
  private _items: IBreadcrumbItem[] = [];
  private _statusMessage: string | undefined;
  private _configurationOpen: boolean = false;
  private _configurationPrompted: boolean = false;
  private _requestSequence: number = 0;

  @override
  public onInit(): Promise<void> {
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    this._locale = LocaleService.getContext(this.context.pageContext);
    this._configurationService = new ConfigurationService(
      this.context.spHttpClient,
      webUrl,
      this.properties.configurationListTitle,
    );
    this._breadcrumbService = new BreadcrumbService(
      this.context.spHttpClient,
      webUrl,
    );

    this.context.placeholderProvider.changedEvent.add(
      this,
      this._onPlaceholdersChanged,
    );
    this.context.application.navigatedEvent.add(this, this._onNavigated);
    this._themeService.start(this._onThemeChanged);
    this._renderPlaceholder();
    this._loadBreadcrumb(true);

    Log.info(LOG_SOURCE, "Dynamic breadcrumb initialized.");
    return Promise.resolve();
  }

  @override
  protected onDispose(): void {
    this._themeService.dispose();
    this.context.placeholderProvider.changedEvent.remove(
      this,
      this._onPlaceholdersChanged,
    );
    this.context.application.navigatedEvent.remove(this, this._onNavigated);
    this._disposePlaceholder();
  }

  private _loadBreadcrumb(forceConfigurationRefresh?: boolean): void {
    const requestSequence: number = ++this._requestSequence;
    this._locale = LocaleService.getContext(this.context.pageContext);

    this._configurationService
      .getConfiguration(forceConfigurationRefresh)
      .then((configuration: IBreadcrumbConfiguration | undefined) => {
        if (requestSequence !== this._requestSequence || this.isDisposed) {
          return undefined;
        }

        this._configuration = configuration;
        if (!configuration) {
          this._items = [];
          this._statusMessage = this._canConfigure()
            ? strings.ConfigurationRequired
            : undefined;
          if (this._canConfigure() && !this._configurationPrompted) {
            this._configurationOpen = true;
            this._configurationPrompted = true;
          }
          this._renderReactComponent();
          return undefined;
        }

        return this._configurationService
          .validateNavigationList(configuration.navigationListId)
          .then(() =>
            this._breadcrumbService.getBreadcrumbForCurrentPage(
              configuration.navigationListId,
              this._locale,
            ),
          )
          .then((items: IBreadcrumbItem[]) => {
            if (requestSequence !== this._requestSequence || this.isDisposed) {
              return;
            }
            this._items = items;
            this._statusMessage = undefined;
            this._renderReactComponent();
          });
      })
      .catch((error: Error) => {
        if (requestSequence !== this._requestSequence || this.isDisposed) {
          return;
        }
        Log.error(LOG_SOURCE, error);
        this._items = [];
        this._statusMessage = this._canConfigure()
          ? strings.ConfigurationError
          : undefined;
        this._renderReactComponent();
      });
  }

  private _renderPlaceholder(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onPlaceholderDisposed },
      );
    }

    if (!this._topPlaceholder) {
      Log.warn(
        LOG_SOURCE,
        "The Top placeholder is not available on this page.",
      );
      return;
    }

    this._renderReactComponent();
  }

  private _renderReactComponent(): void {
    if (!this._topPlaceholder || this.isDisposed) {
      return;
    }

    const element: React.ReactElement<any> = React.createElement(
      DynamicBreadcrumb,
      {
        items: this._items,
        locale: this._locale,
        strings: strings as IDynamicBreadcrumbStrings,
        canConfigure: this._canConfigure(),
        configuration: this._configuration,
        configurationService: this._configurationService,
        configurationOpen: this._configurationOpen,
        statusMessage: this._statusMessage,
        onOpenConfiguration: this._openConfiguration,
        onDismissConfiguration: this._dismissConfiguration,
        onConfigurationSaved: this._configurationSaved,
      },
    );

    ReactDom.render(element, this._topPlaceholder.domElement);
  }

  private _canConfigure(): boolean {
    return this.context.pageContext.web.permissions.hasPermission(
      SPPermission.manageLists,
    );
  }

  private _openConfiguration = (): void => {
    this._configurationOpen = true;
    this._renderReactComponent();
  };

  private _dismissConfiguration = (): void => {
    this._configurationOpen = false;
    this._renderReactComponent();
  };

  private _configurationSaved = (): void => {
    this._configurationOpen = false;
    this._configurationPrompted = true;
    this._loadBreadcrumb(true);
  };

  private _onPlaceholdersChanged = (): void => {
    this._renderPlaceholder();
  };

  private _onNavigated = (): void => {
    this._renderPlaceholder();
    this._loadBreadcrumb();
  };

  private _onThemeChanged = (): void => {
    this._renderReactComponent();
  };

  private _onPlaceholderDisposed = (
    placeholderContent: PlaceholderContent,
  ): void => {
    ReactDom.unmountComponentAtNode(placeholderContent.domElement);
    if (this._topPlaceholder === placeholderContent) {
      this._topPlaceholder = undefined;
    }
  };

  private _disposePlaceholder(): void {
    if (this._topPlaceholder) {
      this._topPlaceholder.dispose();
      this._topPlaceholder = undefined;
    }
  }
}
