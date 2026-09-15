import * as React from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { IContextualMenuItem } from 'office-ui-fabric-react/lib/ContextualMenu';
import { BreadcrumbItem } from './BreadcrumbItem';
import { BreadcrumbConfigurationDialog } from '../dialogs/BreadcrumbConfigurationDialog';
import { IDynamicBreadcrumbProps } from './IDynamicBreadcrumbProps';
import { IBreadcrumbItem } from '../models/IBreadcrumbItem';
import { LocaleService } from '../localization/LocaleService';
import styles from '../styles/DynamicBreadcrumb.module.scss';

/** Modern, responsive and keyboard-accessible breadcrumb presentation. */
export class DynamicBreadcrumb extends React.Component<IDynamicBreadcrumbProps, {}> {
  public render(): JSX.Element {
    const hasItems: boolean = this.props.items.length > 0;
    const showTools: boolean = this.props.canConfigure;

    if (!hasItems && !showTools && !this.props.configurationOpen) {
      return <span className={ styles.empty } />;
    }

    return (
      <div className={ styles.host } dir={ this.props.locale.isRtl ? 'rtl' : 'ltr' }>
        { hasItems ? this._renderBreadcrumb() : this._renderSetupTools() }
        <BreadcrumbConfigurationDialog
          hidden={ !this.props.configurationOpen }
          isRtl={ this.props.locale.isRtl }
          strings={ this.props.strings }
          configuration={ this.props.configuration }
          configurationService={ this.props.configurationService }
          onDismiss={ this.props.onDismissConfiguration }
          onSaved={ this.props.onConfigurationSaved } />
      </div>
    );
  }

  private _renderBreadcrumb(): JSX.Element {
    const items: IBreadcrumbItem[] = this.props.items;
    const segments: JSX.Element[] = [];
    const visibleIndexes: number[] = this._visibleIndexes(items.length);
    visibleIndexes.forEach((itemIndex: number, visibleIndex: number) => {
      if (visibleIndex > 0) {
        const previousIndex: number = visibleIndexes[visibleIndex - 1];
        segments.push(this._separator('separator-before-' + itemIndex));
        if (itemIndex - previousIndex > 1) {
          segments.push(this._overflow(items.slice(previousIndex + 1, itemIndex), 'overflow-' + itemIndex));
          segments.push(this._separator('separator-after-overflow-' + itemIndex));
        }
      }
      segments.push(
        <BreadcrumbItem
          key={ 'item-' + items[itemIndex].id }
          item={ items[itemIndex] }
          isCurrent={ itemIndex === items.length - 1 }
          isRoot={ itemIndex === 0 }
          locale={ this.props.locale } />
      );
    });

    return (
      <nav className={ styles.root } aria-label={ this.props.strings.BreadcrumbAriaLabel }>
        <ol className={ styles.list }>{ segments }</ol>
        { this.props.canConfigure ? this._configurationButton() : undefined }
      </nav>
    );
  }

  private _renderSetupTools(): JSX.Element {
    return (
      <div className={ styles.setupTools }>
        { this.props.statusMessage ?
          <span className={ styles.setupMessage }>{ this.props.statusMessage }</span> : undefined }
        <IconButton
          className={ styles.configureButton }
          iconProps={ { iconName: 'Settings' } }
          ariaLabel={ this.props.strings.ConfigureBreadcrumb }
          title={ this.props.strings.ConfigureBreadcrumb }
          onClick={ this.props.onOpenConfiguration } />
      </div>
    );
  }

  private _configurationButton(): JSX.Element {
    return (
      <IconButton
        className={ styles.configureButton }
        iconProps={ { iconName: 'Settings' } }
        ariaLabel={ this.props.strings.ConfigureBreadcrumb }
        title={ this.props.strings.ConfigureBreadcrumb }
        onClick={ this.props.onOpenConfiguration } />
    );
  }

  private _visibleIndexes(count: number): number[] {
    if (count <= 4) {
      const all: number[] = [];
      for (let index: number = 0; index < count; index++) {
        all.push(index);
      }
      return all;
    }
    return [0, count - 2, count - 1];
  }

  private _separator(key: string): JSX.Element {
    return (
      <li key={ key } className={ styles.separator } aria-hidden={ true }>
        <Icon iconName={ this.props.locale.isRtl ? 'ChevronLeft' : 'ChevronRight' } />
      </li>
    );
  }

  private _overflow(items: IBreadcrumbItem[], key: string): JSX.Element {
    const menuItems: IContextualMenuItem[] = items.map((item: IBreadcrumbItem) => ({
      key: 'overflow-item-' + item.id,
      name: LocaleService.getItemLabel(item, this.props.locale),
      href: item.url,
      target: item.openInNewTab ? '_blank' : undefined
    }));

    return (
      <li key={ key } className={ styles.overflow }>
        <IconButton
          iconProps={ { iconName: 'More' } }
          ariaLabel={ this.props.strings.MoreBreadcrumbItems }
          title={ this.props.strings.MoreBreadcrumbItems }
          menuProps={ { items: menuItems } } />
      </li>
    );
  }
}
