/**
 * SPFx 1.4 does not expose the newer public ThemeProvider. The component stylesheet
 * uses supported SharePoint theme tokens; this observer re-renders Fabric controls
 * when the host updates its theme-related DOM attributes during in-place navigation.
 */
export class ThemeService {
  private _observer: MutationObserver | undefined;

  public start(onThemeChanged: () => void): void {
    if (this._observer || typeof MutationObserver === 'undefined') {
      return;
    }

    this._observer = new MutationObserver((): void => {
      onThemeChanged();
    });

    this._observer.observe(document.documentElement, {
      attributes: true,
      attributeFilter: ['class', 'style']
    });

    if (document.body) {
      this._observer.observe(document.body, {
        attributes: true,
        attributeFilter: ['class', 'style']
      });
    }
  }

  public dispose(): void {
    if (this._observer) {
      this._observer.disconnect();
      this._observer = undefined;
    }
  }
}
