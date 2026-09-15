import { PageContext } from '@microsoft/sp-page-context';
import { IBreadcrumbItem } from '../models/IBreadcrumbItem';

export interface ILocaleContext {
  cultureName: string;
  isRtl: boolean;
  isArabic: boolean;
}

/**
 * Keeps language decisions in one place. PageContext cultureInfo is a supported SPFx
 * API and avoids depending on the legacy _spPageContextInfo global.
 */
export class LocaleService {
  public static getContext(pageContext: PageContext): ILocaleContext {
    const cultureName: string = pageContext.cultureInfo.currentUICultureName ||
      pageContext.cultureInfo.currentCultureName || 'en-us';
    const isArabic: boolean = cultureName.toLowerCase().indexOf('ar') === 0;

    return {
      cultureName: cultureName,
      isRtl: pageContext.cultureInfo.isRightToLeft || isArabic,
      isArabic: isArabic
    };
  }

  public static getItemLabel(item: IBreadcrumbItem, locale: ILocaleContext): string {
    if (locale.isArabic) {
      return item.titleAR || item.title || item.titleEN;
    }

    return item.titleEN || item.title || item.titleAR;
  }
}
