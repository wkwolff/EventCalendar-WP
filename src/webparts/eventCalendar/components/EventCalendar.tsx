/**
 * @file EventCalendar.tsx
 * @description Root React component for the Event Calendar web part. Orchestrates
 *   view switching (calendar vs. list), event data fetching via the `useEvents` hook,
 *   loading/error states, and the event detail side panel. Renders a configuration
 *   placeholder when no list has been selected.
 * @author W. Kevin Wolff
 * @copyright TidalHealth
 */

import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './EventCalendar.module.scss';
import type { IEventCalendarProps } from './IEventCalendarProps';
import { DefaultView } from '../models/IWebPartProps';
import { IEventItem } from '../models/IEventItem';
import { useEvents } from '../hooks/useEvents';
import ViewToggle from './ViewToggle';
import CalendarView from './CalendarView';
import ListView from './ListView';
import FilmstripView from './FilmstripView';
import EventDetailPanel from './EventDetailPanel';

/**
 * Main presentational component for the Event Calendar web part.
 * Manages which view is active (calendar or list), delegates data fetching
 * to the `useEvents` hook, and opens a detail panel when an event is clicked.
 *
 * @param props - Web part configuration passed down from EventCalendarWebPart.
 * @returns The rendered event calendar UI, or a configuration prompt if no list is set.
 */
const EventCalendar: React.FC<IEventCalendarProps> = (props) => {
  const {
    listId,
    fieldMapping,
    selectedFields,
    availableFields,
    viewMode,
    defaultView,
    calendarViewType,
    listLayout,
    maxEvents,
    cardDisplay,
    hasTeamsContext,
  } = props;

  // Determine the initial view: when both views are enabled, respect the user's default;
  // otherwise lock to whichever single view is configured.
  const initialView: DefaultView = viewMode === 'both' ? defaultView : viewMode;
  const [currentView, setCurrentView] = React.useState<DefaultView>(initialView);
  const [selectedEvent, setSelectedEvent] = React.useState<IEventItem | undefined>(undefined);

  // Fetch event data from the SharePoint list using the configured field mapping
  const { events, loading, error } = useEvents(listId, fieldMapping, selectedFields, maxEvents);

  // Sync the active view whenever the property pane settings change
  React.useEffect(() => {
    if (viewMode === 'both') {
      setCurrentView(defaultView);
    } else {
      setCurrentView(viewMode);
    }
  }, [viewMode, defaultView]);

  // Configuration placeholder — shown when no list has been selected yet
  if (!listId) {
    return (
      <section className={`${styles.eventCalendar} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.configure}>
          <div className={styles.configureIcon}>
            <Icon iconName="Calendar" />
          </div>
          <div>Select a list in the web part settings to display events.</div>
        </div>
      </section>
    );
  }

  /**
   * Renders the appropriate list-style view based on the `listLayout` property.
   * @returns Either a FilmstripView or a compact ListView component.
   */
  const renderListView = (): React.ReactElement => {
    if (listLayout === 'filmstrip') {
      return (
        <FilmstripView
          events={events}
          availableFields={availableFields}
          cardDisplay={cardDisplay}
          onEventClick={setSelectedEvent}
        />
      );
    }
    return (
      <ListView
        events={events}
        availableFields={availableFields}
        cardDisplay={cardDisplay}
        onEventClick={setSelectedEvent}
      />
    );
  };

  return (
    <section className={`${styles.eventCalendar} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.header}>
        <span className={styles.title}>Events</span>
        {/* View toggle only appears when both calendar and list views are enabled */}
        {viewMode === 'both' && (
          <ViewToggle currentView={currentView} onViewChange={setCurrentView} />
        )}
      </div>

      {error && <div className={styles.error}>{error}</div>}

      {loading ? (
        <Spinner size={SpinnerSize.large} label="Loading events..." />
      ) : currentView === 'calendar' ? (
        <CalendarView
          events={events}
          viewType={calendarViewType}
          onEventClick={setSelectedEvent}
        />
      ) : (
        renderListView()
      )}

      {/* Side panel slides in when an event is clicked from any view */}
      <EventDetailPanel
        event={selectedEvent}
        availableFields={availableFields}
        isOpen={!!selectedEvent}
        onDismiss={() => setSelectedEvent(undefined)}
      />
    </section>
  );
};

export default EventCalendar;
