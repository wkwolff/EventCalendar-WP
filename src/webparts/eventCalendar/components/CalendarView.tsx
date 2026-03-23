/**
 * @file CalendarView.tsx
 * @description Wraps the FullCalendar library to render events in month, week, or day
 *   grid views. Maps internal `IEventItem` objects to FullCalendar's `EventInput`
 *   format and routes click events back up to open the detail panel.
 * @author W. Kevin Wolff
 * @copyright Wolff Creative
 */

import * as React from 'react';
import FullCalendar from '@fullcalendar/react';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import interactionPlugin from '@fullcalendar/interaction';
import { EventInput } from '@fullcalendar/core';
import { IEventItem } from '../models/IEventItem';
import { CalendarViewType } from '../models/IWebPartProps';
import styles from './CalendarView.module.scss';

/**
 * Props for the CalendarView component.
 */
export interface ICalendarViewProps {
  /** Array of event items to display on the calendar. */
  events: IEventItem[];
  /** FullCalendar initial view: 'dayGridMonth', 'timeGridWeek', or 'timeGridDay'. */
  viewType: CalendarViewType;
  /** Callback invoked when a calendar event is clicked. */
  onEventClick: (event: IEventItem) => void;
}

/**
 * Renders a FullCalendar instance with month/week/day navigation.
 * Each `IEventItem` is stored in `extendedProps.eventItem` so it can be
 * retrieved on click without a secondary lookup.
 *
 * @param props - Calendar view configuration and event data.
 * @returns A FullCalendar component wrapped in a styled container.
 */
const CalendarView: React.FC<ICalendarViewProps> = ({ events, viewType, onEventClick }) => {
  // Transform internal event items into FullCalendar's EventInput format.
  // The original IEventItem is stashed in extendedProps for retrieval on click.
  const calendarEvents: EventInput[] = events.map(e => ({
    id: String(e.id),
    title: e.title,
    start: e.startDate,
    end: e.endDate,
    allDay: e.allDay,
    extendedProps: { eventItem: e },
  }));

  return (
    <div className={styles.calendarView}>
      <FullCalendar
        plugins={[dayGridPlugin, timeGridPlugin, interactionPlugin]}
        initialView={viewType}
        headerToolbar={{
          left: 'prev,next today',
          center: 'title',
          right: 'dayGridMonth,timeGridWeek,timeGridDay',
        }}
        events={calendarEvents}
        eventClick={(info) => {
          // Retrieve the stashed IEventItem from FullCalendar's extendedProps
          const item = info.event.extendedProps.eventItem as IEventItem;
          onEventClick(item);
        }}
        height="auto"
        nowIndicator
        dayMaxEvents={3}
      />
    </div>
  );
};

export default CalendarView;
