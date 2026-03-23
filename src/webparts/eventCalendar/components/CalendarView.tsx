import * as React from 'react';
import FullCalendar from '@fullcalendar/react';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import interactionPlugin from '@fullcalendar/interaction';
import { EventInput } from '@fullcalendar/core';
import { IEventItem } from '../models/IEventItem';
import { CalendarViewType } from '../models/IWebPartProps';
import styles from './CalendarView.module.scss';

export interface ICalendarViewProps {
  events: IEventItem[];
  viewType: CalendarViewType;
  onEventClick: (event: IEventItem) => void;
}

const CalendarView: React.FC<ICalendarViewProps> = ({ events, viewType, onEventClick }) => {
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
