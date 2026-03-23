/**
 * @file IWebPartProps.ts
 * @description Property bag interfaces and type aliases for the Event Calendar
 *              web part. These types are consumed by the SPFx property pane and
 *              passed down to the root React component.
 * @author W. Kevin Wolff
 * @copyright TidalHealth
 */

/** FullCalendar initial view variants exposed to end users. */
export type CalendarViewType = 'dayGridMonth' | 'timeGridWeek' | 'timeGridDay';

/** Which panel is shown first when both calendar and list are enabled. */
export type DefaultView = 'calendar' | 'list';

/** Controls whether the calendar, list, or both panels are rendered. */
export type ViewMode = 'both' | 'calendar' | 'list';

/** Layout style for the list/card view. */
export type ListLayout = 'filmstrip' | 'compact';

/**
 * Serializable property bag persisted by SPFx for this web part instance.
 *
 * Column-mapping fields (`titleField`, `startDateField`, etc.) store the
 * **internal name** of the SharePoint column so queries work regardless of
 * display-name changes.
 */
export interface IEventCalendarWebPartProps {
  /** GUID of the selected SharePoint list. */
  selectedListId: string;

  /** Internal name of the column mapped to the event title. */
  titleField: string;

  /** Internal name of the DateTime column used as event start. */
  startDateField: string;

  /** Internal name of the DateTime column used as event end (may be empty). */
  endDateField: string;

  /** Internal name of the Boolean column indicating an all-day event (may be empty). */
  allDayField: string;

  /** Internal name of the Choice/Text column used for category labels (may be empty). */
  categoryField: string;

  /** Internal name of the Text column used for event location (may be empty). */
  locationField: string;

  /** Internal names of additional columns to fetch and display in event details. */
  selectedFields: string[];

  /** Whether to show calendar only, list only, or both panels. */
  viewMode: ViewMode;

  /** Which panel is active by default when `viewMode` is `'both'`. */
  defaultView: DefaultView;

  /** Initial FullCalendar view (month, week, or day). */
  calendarViewType: CalendarViewType;

  /** Card layout style for the list panel. */
  listLayout: ListLayout;

  /** Maximum number of events to retrieve from SharePoint. */
  maxEvents: number;

  /** Show the category badge on event cards. */
  showCategory: boolean;

  /** Show the location line on event cards. */
  showLocation: boolean;

  /** Show the time range on event cards. */
  showTime: boolean;

  /** Show the thumbnail image on event cards (when available). */
  showImage: boolean;
}
