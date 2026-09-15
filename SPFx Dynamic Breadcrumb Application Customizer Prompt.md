Act as a Senior SharePoint Framework (SPFx) Architect and Developer with strong experience in **SharePoint Subscription Edition On-Premises**, React, TypeScript, Fluent UI, SharePoint REST APIs, localization, theming, and SPFx Extensions.

I want you to design and implement a production-ready **custom dynamic Breadcrumb component using an SPFx Application Customizer Extension** for **SharePoint Subscription Edition running on a local/on-premises SharePoint farm**.

The solution must follow SharePoint Subscription Edition compatibility requirements and must use an SPFx version, React version, TypeScript version, Fluent UI version, Node.js version, and build tooling that are supported by SharePoint Subscription Edition.

Do not design this as a normal SPFx Web Part. The breadcrumb must be implemented as an **Application Customizer** so that it can be rendered globally across the SharePoint site.

The solution must satisfy the following requirements.

### 1. SPFx Application Customizer

Create a custom SPFx **Application Customizer Extension** responsible for rendering a reusable breadcrumb/navigation component.

The breadcrumb should be capable of appearing automatically across modern SharePoint pages after the extension is deployed and activated.

Use the appropriate SharePoint placeholders, such as `PlaceholderName.Top`, when suitable for the design.

The implementation must be compatible with:

- SharePoint Subscription Edition
- On-premises SharePoint deployment
- Modern SharePoint pages
- SPFx Extensions
- React
- TypeScript
- Fluent UI supported by the selected SPFx version

Avoid APIs or features that are available only in SharePoint Online if they are not supported by SharePoint Subscription Edition.

---

### 2. Dynamic Configuration

The breadcrumb must be completely configurable.

Before the breadcrumb starts retrieving its navigation information, administrators must have a configuration mechanism that allows them to define the SharePoint List that will act as the breadcrumb's data source.

The configuration interface should allow the administrator to:

- Select an existing SharePoint List.
- Specify the list name or list ID.
- Validate that the selected list exists.
- Validate that the required columns exist.
- Save the selected configuration.
- Change the configuration later.
- Reconfigure the breadcrumb without rebuilding or redeploying the SPFx package.

If the required SharePoint List does not exist, provide an option such as:

**Create Breadcrumb Configuration List**

The solution should then automatically create the required SharePoint List and fields.

Because this is an **Application Customizer rather than a Web Part**, design an appropriate configuration experience for an SPFx Extension.

Possible approaches may include:

- A dedicated configuration page.
- A SharePoint List-based configuration screen.
- An SPFx configuration dialog.
- Extension properties combined with a configuration list.
- Site-scoped configuration.

Recommend the best architecture for SharePoint Subscription Edition.

The generated configuration must persist after page reloads.

---

### 3. Dynamic SharePoint List Data Source

The breadcrumb items must be retrieved dynamically from a SharePoint List.

Design an appropriate list schema supporting hierarchical navigation.

For example, consider fields such as:

- Title
- TitleEN
- TitleAR
- URL
- ParentId
- DisplayOrder
- IsActive
- OpenInNewTab
- Icon
- Audience or permissions if applicable

The breadcrumb should dynamically determine the correct hierarchy.

For example:

Home > Department > Services > Current Page

Do not hardcode breadcrumb entries.

The breadcrumb must refresh according to configuration and list content.

Use a SharePoint-compatible API approach such as the SPFx `SPHttpClient` or another API supported by SharePoint Subscription Edition.

---

### 4. English / Arabic Localization

The breadcrumb must fully support multilingual SharePoint environments, with particular focus on:

- English
- Arabic

The component must automatically detect the currently selected SharePoint UI language.

When English is selected:

- Use English labels.
- Use Left-to-Right layout.
- Apply:

```css
direction: ltr;
```

When Arabic is selected:

- Use Arabic labels.
- Use Right-to-Left layout.
- Apply:

```css
direction: rtl;
```

The switch between languages must affect:

- Breadcrumb labels
- Navigation direction
- Chevron/separator direction
- Alignment
- Icons where necessary
- Text
- Accessibility labels
- Layout

For example:

English:

Home > Services > Cloud Solutions

Arabic:

الرئيسية < الخدمات < الحلول السحابية

The component should obtain localized values from fields such as:

- `TitleEN`
- `TitleAR`

or another recommended multilingual architecture.

Also use the SPFx localization mechanism where appropriate using locale resource files such as:

```text
en-us.js
ar-sa.js
```

Explain how SharePoint UI culture can be detected inside an SPFx Application Customizer.

---

### 5. Fully Responsive Design

The breadcrumb must be fully responsive.

It should display correctly on:

- Desktop
- Laptop
- Tablet
- Mobile

The component must adapt gracefully when the breadcrumb contains many levels.

Consider responsive techniques such as:

- Flexbox
- wrapping
- truncation
- ellipsis
- collapsed breadcrumb items
- responsive breakpoints
- Fluent UI responsive utilities where supported

For example, a long desktop breadcrumb:

Home > Corporate > Departments > Technology > Infrastructure > Networks

could become something similar to:

Home > ... > Infrastructure > Networks

on smaller devices.

The solution must remain usable and accessible at all screen sizes.

---

### 6. Dynamic SharePoint Theme Integration

The breadcrumb must automatically inherit and react to the current SharePoint site theme.

Do not hardcode the main branding colors.

When a SharePoint administrator changes the site's theme, the breadcrumb styling should automatically adapt.

The component should dynamically use appropriate SharePoint theme values for elements such as:

- Primary color
- Text color
- Background color
- Link color
- Hover color
- Border color
- Separator color
- Focus color

Use the SPFx theming APIs supported by SharePoint Subscription Edition.

Handle theme changes without requiring the SPFx package to be rebuilt.

Explain how the Application Customizer can subscribe to SharePoint theme changes and update React component styles dynamically.

---

### 7. Modern SharePoint Look and Feel

The breadcrumb design must closely follow the **Modern SharePoint user experience**.

Use Fluent UI components and design principles where supported by the SharePoint Subscription Edition SPFx version.

The breadcrumb should visually look like it belongs natively inside SharePoint rather than appearing as a third-party component.

Follow SharePoint/Fluent UI design principles for:

- Typography
- Font sizes
- Spacing
- Padding
- Hover states
- Focus states
- Colors
- Separators
- Icons
- Accessibility
- Responsive behavior

The component should be clean, minimal, professional, enterprise-ready, and visually consistent with Microsoft SharePoint.

---

### 8. Recommended Architecture

Use a clean project structure similar to:

```text
src/
  extensions/
    dynamicBreadcrumb/
      DynamicBreadcrumbApplicationCustomizer.ts

      components/
        DynamicBreadcrumb.tsx
        IDynamicBreadcrumbProps.ts
        BreadcrumbItem.tsx

      services/
        BreadcrumbService.ts
        ConfigurationService.ts

      models/
        IBreadcrumbItem.ts
        IBreadcrumbConfiguration.ts

      localization/
        LocaleService.ts

      theming/
        ThemeService.ts

      dialogs/
        BreadcrumbConfigurationDialog.tsx

      styles/
        DynamicBreadcrumb.module.scss
```

You may improve this structure if there is a better architecture.

Use proper separation of concerns between:

- UI
- SharePoint communication
- configuration
- localization
- theming
- data models
- business logic

---

### 9. Configuration/List Provisioning

Provide a mechanism that checks whether the required breadcrumb list exists.

Conceptually:

```text
Application Customizer starts
        |
        v
Read breadcrumb configuration
        |
        +---- Configuration exists ----+
        |                              |
        v                              v
Load configured list             No configuration
        |                              |
        v                              v
Load breadcrumb data       Open/offer configuration
                                       |
                                       v
                           Select existing SharePoint List
                                       OR
                                       |
                                       v
                             Create required list
                                       |
                                       v
                              Save configuration
                                       |
                                       v
                              Render breadcrumb
```

The list creation functionality should safely:

1. Check whether the list already exists.
2. Create the list if needed.
3. Create the required fields.
4. Avoid recreating existing fields.
5. Optionally add default navigation records.
6. Return useful errors if provisioning fails.

---

### 10. Error Handling

Implement professional error handling.

The breadcrumb must handle cases such as:

- Configuration list missing.
- Navigation list missing.
- Required column missing.
- REST request failure.
- Invalid URL.
- Empty breadcrumb.
- User without list access.
- SharePoint language unavailable.
- Theme information unavailable.

The extension must not break the SharePoint page if an error occurs.

Use appropriate logging and graceful fallback behavior.

---

### 11. Performance

Optimize the component for enterprise SharePoint environments.

Consider:

- Avoiding unnecessary REST requests.
- Caching configuration.
- Caching navigation data when appropriate.
- Minimizing React re-renders.
- Loading only required fields.
- Using `$select`, `$filter`, and `$orderby`.
- Cleaning up event handlers when the Application Customizer is disposed.

---

### 12. Accessibility

Follow accessibility best practices.

The breadcrumb should use appropriate semantic markup such as:

```html
<nav aria-label="Breadcrumb">
```

and appropriate ARIA attributes.

Support:

- Keyboard navigation
- Screen readers
- Focus states
- Color contrast
- RTL accessibility

---

### 13. Expected Deliverables

Generate the complete implementation step by step.

Start by explaining the architecture and the design decisions.

Then provide:

1. Exact SPFx version recommended for SharePoint Subscription Edition.
2. Compatible Node.js version.
3. Compatible React version.
4. Compatible TypeScript version.
5. Compatible Fluent UI version.
6. SPFx project creation commands.
7. Application Customizer creation commands.
8. Complete folder structure.
9. Interfaces and models.
10. SharePoint List schema.
11. SharePoint List creation/provisioning service.
12. Configuration service.
13. Configuration UI.
14. Breadcrumb data service.
15. Localization implementation.
16. Arabic RTL handling.
17. SharePoint theme integration.
18. Responsive SCSS.
19. React breadcrumb component.
20. Application Customizer implementation.
21. Placeholder rendering.
22. Error handling.
23. Caching/performance implementation.
24. Package configuration.
25. Feature XML or ClientSideInstance configuration if required.
26. `.sppkg` generation steps.
27. Deployment to the SharePoint Subscription Edition App Catalog.
28. Activation/configuration on a SharePoint site.
29. Testing instructions.
30. Troubleshooting instructions.

For every important file, provide its **complete source code**, not only partial snippets.

Do not use SharePoint Online-only functionality unless you explicitly identify it and provide an equivalent solution supported by SharePoint Subscription Edition.

The final solution should be:

**Dynamic + Configurable + Localized + RTL/LTR aware + Responsive + Theme-aware + Modern SharePoint styled + Reusable + Maintainable + Production-ready.**