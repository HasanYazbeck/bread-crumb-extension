# Dynamic Breadcrumb Application Customizer

An SPFx Application Customizer for SharePoint Server Subscription Edition (SE). It places a configurable, list-driven breadcrumb in the modern page `Top` placeholder, so the experience is available site-wide rather than being added to individual pages.

## Architecture

```text
Application Customizer
        │
        ├── Breadcrumb Configuration list (site-scoped selection)
        │        └── Navigation list ID and title
        │
        ├── Breadcrumb Navigation list (hierarchical, managed content)
        │        └── active items → current URL match → parent chain
        │
        └── React + Office UI Fabric React view in PlaceholderName.Top
                 ├── responsive overflow menu
                 ├── English / Arabic labels and LTR / RTL layout
                 └── SharePoint theme-token styles
```

The customizer uses only SPFx `SPHttpClient`, `PageContext`, placeholders, and SharePoint REST endpoints. It does not depend on Microsoft Graph or SharePoint Online-only services.

## Compatibility

This project intentionally targets the broad SPSE baseline already present in the solution:

| Component | Version |
| --- | --- |
| SPFx | 1.4.1 |
| Node.js | 8.17.0 (tested) |
| TypeScript | 2.4.2 |
| React / ReactDOM | 15.6.2 |
| Office UI Fabric React (Fluent UI predecessor) | 5.21.0 |

SPFx 1.4.1 works on early SharePoint SE farms; farms patched to the SE 23H1 feature update can also support SPFx 1.5.1. Do not move this package to a newer SPFx version unless the target farm explicitly supports it—the server and package versions must match.

## Configuration experience

Users with **Manage Lists** permission see a settings button. On the first unconfigured visit, they are offered the configuration dialog automatically. The dialog can:

- choose an existing generic SharePoint list;
- accept its exact name or GUID;
- validate its required fields before saving;
- create a navigation list and all required fields safely; and
- create the configuration list on the first successful save.

The persisted configuration is in **Breadcrumb Configuration** (change the `configurationListTitle` custom-action property if a different name is needed). This makes later reconfiguration possible without a rebuild or redeployment.

### Navigation list schema

| Internal name | Type | Purpose |
| --- | --- | --- |
| `Title` | Single line of text | Fallback label |
| `TitleEN` | Single line of text | English label |
| `TitleAR` | Single line of text | Arabic label |
| `Url` | Hyperlink | Destination and URL matching key |
| `ParentId` | Number | Parent item ID; blank for a root node |
| `DisplayOrder` | Number | Sort order among siblings |
| `IsActive` | Yes/No | Includes/excludes the item |
| `OpenInNewTab` | Yes/No | Adds `_blank` with safe `rel` attributes |
| `Icon` | Single line of text | Optional Office UI Fabric icon name |

Create a root item such as Home, then point each child `ParentId` to its parent record. The current page is selected by the longest matching active `Url`, then the component walks back through its parents. This avoids hardcoded navigation and detects parent loops safely.

## Localization, theme, and accessibility

`PageContext.cultureInfo.currentUICultureName` selects `TitleAR` or `TitleEN`; Arabic UI culture (or a right-to-left UI culture) also sets `dir="rtl"` and reverses chevrons. `en-us.js` and `ar-sa.js` localize all component controls.

The SCSS uses SharePoint theme tokens instead of fixed branding colors. On this SPFx baseline, the public newer `ThemeProvider` API is unavailable; theme-token styles are therefore the supported mechanism. A lightweight DOM observer rerenders Fabric controls when the page host updates theme attributes.

The view is a semantic `<nav aria-label="Breadcrumb">` containing an ordered list. Long trails collapse middle items into a keyboard-accessible Fabric overflow menu. Labels truncate cleanly and the layout wraps at narrow widths.

## Build and package

Use the Node version stated above, then run:

```bash
npm install
npm run build
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is written to `sharepoint/solution/breadcrumb-extension.sppkg`.

## Deploy to SharePoint SE

1. Upload the `.sppkg` to the on-premises App Catalog and trust the package.
2. Install the app on the target site collection.
3. Activate the solution feature, which installs the `ClientSideExtension.ApplicationCustomizer` custom action from `sharepoint/assets/elements.xml`.
4. Open a modern page as a user with **Manage Lists** permission, use the configuration dialog, and either select a validated list or choose **Create breadcrumb navigation list**.
5. Add navigation records, populate both `TitleEN` and `TitleAR`, set their `Url` and parent IDs, then refresh a matching page.

For a tenant/site deployment model where automatic feature activation is not desired, add the same Application Customizer custom action manually with component ID `dbf24974-a827-4696-a647-b180de8c712a` and optional properties:

```json
{
  "configurationListTitle": "Breadcrumb Configuration"
}
```

## Verification checklist

- Visit a configured page in English and Arabic; verify labels, text direction, separators, and settings UI direction.
- Test a deep hierarchy on desktop and a narrow mobile viewport; the overflow menu must expose skipped levels.
- Switch the site theme and navigate to another modern page; links, focus indication, borders, and backgrounds should inherit the new theme.
- Remove a required field or deny list access with a test account; the page remains usable, while administrators see a safe configuration prompt and details are logged to the SPFx developer dashboard.
- Change the selected navigation list, save, and navigate again; the persisted configuration should be used without rebuilding the package.

## Troubleshooting

- **No breadcrumb:** ensure the page is modern, the Top placeholder is available, and an active item URL matches the page path.
- **Configuration cannot save:** the administrator needs Manage Lists permission and the current user needs access to the selected navigation list.
- **Validation failure:** restore the exact internal field names in the schema table; display names alone are not sufficient.
- **Theme does not match:** confirm the page is using a SharePoint theme and avoid overriding component CSS with hardcoded site styles.
- **Package is rejected:** verify the SPFx version against the exact SE farm patch level before deployment.
