# TH - Events

A SharePoint Framework (SPFx) web part that displays events from any SharePoint list in filmstrip, compact list, or full calendar views. Designed for TidalHealth SharePoint Online sites.

![SPFx Version](https://img.shields.io/badge/SPFx-1.22.2-green.svg)
![Node.js Version](https://img.shields.io/badge/Node.js-v22.14+-blue.svg)

## Features

- **Filmstrip layout** — Horizontal scrolling event cards with pagination dots, matching the OOB SharePoint Events web part design
- **Compact list layout** — Vertical event list grouped by relative dates (Today, Tomorrow, This week, etc.)
- **Calendar view** — Full month/week/day calendar powered by FullCalendar
- **Flexible column mapping** — Map any list's columns to title, start date, end date, all-day, category, and location fields
- **Additional field display** — Select extra columns to show on cards and in the detail panel
- **Card display toggles** — Show/hide image, category, time, and location from the property pane
- **Auto-detection** — Automatically detects standard Events list fields and common column names on list selection
- **Auto-linking** — Email addresses, URLs, and SharePoint file paths are automatically rendered as clickable links
- **Theme-aware** — Inherits SharePoint site theme colors for a native look and feel
- **Multi-line text support** — Rich text and multi-line fields render cleanly with line clamping
- **Detail panel** — Click any event to open a side panel with full event details, images, and field values
- **Teams compatible** — Works in SharePoint pages, Teams tabs, and Teams personal apps

## Solution

| Solution | Author(s) |
|----------|-----------|
| TH - Events | W. Kevin Wolff, Microsoft Developer, TidalHealth |

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0.0 | March 23, 2026 | Initial production release |

## Prerequisites

- [Node.js](https://nodejs.org/) v22.14.0+ (LTS)
- SharePoint Online environment with App Catalog
- SharePoint Framework v1.22.2

## Getting Started

### 1. Install dependencies

```bash
npm install
```

### 2. Run locally (development)

```bash
npm run start
```

This starts the local workbench dev server at `https://localhost:4321` and opens your SharePoint site's workbench. Add the **TH - Events** web part to the page and configure it via the property pane.

### 3. Build for production

```bash
npm run build
```

The `.sppkg` package is generated at:

```
sharepoint/solution/event-calendar-wp.sppkg
```

### 4. Clean build artifacts

```bash
npm run clean
```

## Deployment

1. Run `npm run build` to create the production package
2. Navigate to your tenant **App Catalog** site (e.g., `https://yourtenant.sharepoint.com/sites/appcatalog`)
3. Upload `sharepoint/solution/event-calendar-wp.sppkg` to the **Apps for SharePoint** library
4. In the deployment dialog, check **"Make this solution available to all sites in the organization"** for tenant-wide deployment
5. On any SharePoint page, edit the page and add the **TH - Events** web part

## Configuration

### Property Pane — Page 1: Data Source

| Setting | Description |
|---------|-------------|
| **Select a list** | Pick any SharePoint list containing event data |
| **Title column** | Column used for the event title (auto-detected) |
| **Start date column** | Column for event start date/time (auto-detected) |
| **End date column** | Column for event end date/time (auto-detected) |
| **All day column** | Boolean column indicating all-day events (optional) |
| **Category column** | Choice/text column for event category labels (optional) |
| **Location column** | Text column for event location/venue (optional) |
| **Additional Fields** | Check any extra columns to display on cards and in the detail panel |

### Property Pane — Page 2: Display Settings

| Setting | Description |
|---------|-------------|
| **Available Views** | Both (Calendar & List), Calendar Only, or List Only |
| **Default View** | Which view loads first when both are enabled |
| **Calendar View** | Month, Week, or Day for the calendar default |
| **List Layout** | Filmstrip (horizontal cards) or Compact list (grouped rows) |
| **Maximum Events** | Number of events to fetch (10–200) |
| **Show image** | Toggle event image/thumbnail display on cards |
| **Show category** | Toggle category label display on cards |
| **Show date and time** | Toggle date/time display on cards |
| **Show location** | Toggle location display on cards |

## Project Structure

```
src/webparts/eventCalendar/
  EventCalendarWebPart.ts             # Web part entry point and property pane config
  EventCalendarWebPart.manifest.json  # Web part manifest (ID, title, hosts)
  components/
    EventCalendar.tsx                  # Root component — view switching, loading, errors
    IEventCalendarProps.ts             # Component prop and display option interfaces
    CalendarView.tsx                   # FullCalendar month/week/day wrapper
    FilmstripView.tsx                  # Horizontal scrolling card carousel with pagination
    ListView.tsx                       # Grouped compact event list with relative dates
    EventCard.tsx                      # Event card renderer (filmstrip + compact variants)
    EventDetailPanel.tsx               # Fluent UI side panel with full event details
    ViewToggle.tsx                     # Calendar/List icon toggle buttons
    FieldBadge.tsx                     # Dynamic field value renderer (links, text, images)
    *.module.scss                      # Component-scoped SCSS styles with theme tokens
  models/
    IEventItem.ts                      # Normalized event data model
    IFieldInfo.ts                      # SharePoint field metadata model
    IWebPartProps.ts                   # Web part property types and layout enums
  services/
    EventService.ts                    # SharePoint list item fetching and field mapping
    FieldService.ts                    # Field metadata fetching and auto-detection logic
    PnPSetup.ts                        # PnPjs SPFx context initialization
  hooks/
    useEvents.ts                       # React hook — fetches and caches event data
    useFields.ts                       # React hook — fetches list field metadata
  loc/
    en-us.js                           # English localization strings
    mystrings.d.ts                     # TypeScript string declarations
```

## Technology Stack

| Technology | Version | Purpose |
|-----------|---------|---------|
| SharePoint Framework | 1.22.2 | SPFx web part platform |
| React | 17.0.1 | UI component rendering |
| TypeScript | 5.8.x | Static type checking |
| Fluent UI React | 8.x | SharePoint-native UI controls (Spinner, Panel, Icon, Pivot) |
| FullCalendar | 6.x | Interactive calendar grid (month, week, day views) |
| PnPjs | 4.x | SharePoint REST API client with type-safe queries |
| Heft | 1.1.2 | Build orchestration and bundling |
| Webpack | 5.x | Module bundling (managed by SPFx build rig) |

## Browser Support

- Microsoft Edge (Chromium)
- Google Chrome
- Mozilla Firefox
- Safari (macOS)

## Known Limitations

- **Lookup/Person expand columns**: Fields requiring `$expand` (e.g., ParticipantsPicker) are excluded from REST queries to prevent API errors. Display the lookup's ID column as a workaround.
- **Image detection**: Only URLs ending in standard image extensions (`.jpg`, `.png`, `.gif`, `.webp`, `.svg`, `.bmp`) are rendered as card thumbnails. Other URL fields are treated as links.
- **Filmstrip card width**: Cards use a fixed 260px width. Responsive card sizing is planned for a future release.
- **No recurrence expansion**: Recurring events are displayed as single items; recurrence patterns are not expanded into individual occurrences.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Web part not appearing in toolbox | Hard-refresh the workbench (Ctrl+Shift+R). If deployed, ensure the app is activated on the site. |
| "Loading fields..." stays indefinitely | Verify the current user has at least Read access to the selected list. |
| Events not loading | Check that Start Date column mapping is set. Open browser DevTools (F12) → Console for error details. |
| Broken images on cards | The image URL column value must end in a standard image extension. Non-image URLs display the theme placeholder instead. |
| Styles look wrong / no theme colors | Ensure `supportsThemeVariants` is `true` in the manifest (it is by default). Clear browser cache if needed. |

## References

- [SharePoint Framework Documentation](https://docs.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
- [PnPjs Documentation](https://pnp.github.io/pnpjs/)
- [FullCalendar React](https://fullcalendar.io/docs/react)
- [Fluent UI React](https://developer.microsoft.com/fluentui)
- [SPFx Property Controls](https://pnp.github.io/sp-dev-fx-property-controls/)

## License

This project is proprietary to TidalHealth. Unauthorized distribution is prohibited.

---

*Developed by W. Kevin Wolff, Microsoft Developer, TidalHealth*
