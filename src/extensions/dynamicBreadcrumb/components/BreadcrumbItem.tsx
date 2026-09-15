import * as React from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { IBreadcrumbItem } from '../models/IBreadcrumbItem';
import { ILocaleContext, LocaleService } from '../localization/LocaleService';
import styles from '../styles/DynamicBreadcrumb.module.scss';

export interface IBreadcrumbItemComponentProps {
  item: IBreadcrumbItem;
  isCurrent: boolean;
  locale: ILocaleContext;
}

/** A single, semantic breadcrumb node. */
export function BreadcrumbItem(props: IBreadcrumbItemComponentProps): JSX.Element {
  const label: string = LocaleService.getItemLabel(props.item, props.locale);
  const icon: JSX.Element | undefined = props.item.icon ?
    <Icon aria-hidden={ true } className={ styles.itemIcon } iconName={ props.item.icon } /> : undefined;
  const target: string | undefined = props.item.openInNewTab ? '_blank' : undefined;

  return (
    <li className={ styles.item }>
      { props.isCurrent ?
        <span className={ styles.currentItem } aria-current='page' title={ label }>
          { icon }<span>{ label }</span>
        </span> :
        <a className={ styles.link } href={ props.item.url } target={ target }
          rel={ props.item.openInNewTab ? 'noopener noreferrer' : undefined } title={ label }>
          { icon }<span>{ label }</span>
        </a>
      }
    </li>
  );
}
