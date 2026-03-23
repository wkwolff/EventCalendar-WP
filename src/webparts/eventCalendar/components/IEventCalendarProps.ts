/**
 * @file IEventCalendarProps.ts
 * @description TypeScript interfaces for the EventCalendar root component props
 *   and card display configuration. These interfaces bridge the SPFx web part
 *   property bag and the React component tree.
 * @author W. Kevin Wolff
 * @copyright TidalHealth
 */

import { DefaultView, CalendarViewType, ViewMode, ListLayout } from '../models/IWebPartProps';
import { IFieldInfo } from '../models/IFieldInfo';
import { IEventFieldMapping } from '../services/EventService';

/**
 * Controls which optional metadata sections are visible on event cards.
 * Each flag corresponds to a toggle in the property pane's "Card Display" group.
 */
export interface ICardDisplayOptions {
  /** Whether to show the category badge/label on cards. */
  showCategory: boolean;
  /** Whether to show the location line on cards. */
  showLocation: boolean;
  /** Whether to show the time range or "All day" indicator on cards. */
  showTime: boolean;
  /** Whether to show the hero/banner image on cards. */
  showImage: boolean;
}

/**
 * Props passed from the SPFx web part class to the root EventCalendar component.
 * Encapsulates all configuration needed to fetch, display, and interact with events.
 */
export interface IEventCalendarProps {
  /** GUID of the selected SharePoint list that contains event data. */
  listId: string;
  /** Maps core event fields (title, dates, category, etc.) to SharePoint column internal names. */
  fieldMapping: IEventFieldMapping;
  /** Internal names of additional (non-core) fields the user opted to display. */
  selectedFields: string[];
  /** Metadata for all display-eligible fields (excludes fields already mapped to core slots). */
  availableFields: IFieldInfo[];
  /** Which views are enabled: 'both', 'calendar', or 'list'. */
  viewMode: ViewMode;
  /** The initial active view when viewMode is 'both'. */
  defaultView: DefaultView;
  /** FullCalendar initial view type: 'dayGridMonth', 'timeGridWeek', or 'timeGridDay'. */
  calendarViewType: CalendarViewType;
  /** List-mode layout style: 'filmstrip' (horizontal cards) or 'compact' (stacked rows). */
  listLayout: ListLayout;
  /** Maximum number of events to fetch from the SharePoint list. */
  maxEvents: number;
  /** Toggles for optional card metadata sections. */
  cardDisplay: ICardDisplayOptions;
  /** True when the current SharePoint theme is inverted (dark mode). */
  isDarkTheme: boolean;
  /** True when the web part is hosted inside Microsoft Teams. */
  hasTeamsContext: boolean;
}
