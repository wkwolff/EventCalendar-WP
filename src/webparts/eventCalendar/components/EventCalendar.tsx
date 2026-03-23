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

  const initialView: DefaultView = viewMode === 'both' ? defaultView : viewMode;
  const [currentView, setCurrentView] = React.useState<DefaultView>(initialView);
  const [selectedEvent, setSelectedEvent] = React.useState<IEventItem | undefined>(undefined);
  const { events, loading, error } = useEvents(listId, fieldMapping, selectedFields, maxEvents);

  React.useEffect(() => {
    if (viewMode === 'both') {
      setCurrentView(defaultView);
    } else {
      setCurrentView(viewMode);
    }
  }, [viewMode, defaultView]);

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
