import * as React from 'react';
import { IEventItem } from '../models/IEventItem';
import { IFieldInfo } from '../models/IFieldInfo';
import { ICardDisplayOptions } from './IEventCalendarProps';
import EventCard from './EventCard';
import styles from './ListView.module.scss';

export interface IListViewProps {
  events: IEventItem[];
  availableFields: IFieldInfo[];
  cardDisplay: ICardDisplayOptions;
  onEventClick: (event: IEventItem) => void;
}

interface IGroupedEvents {
  label: string;
  events: IEventItem[];
}

function getRelativeDateLabel(date: Date, now: Date): string {
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const target = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const diffDays = Math.round((target.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));

  if (diffDays < 0) return 'Earlier';
  if (diffDays === 0) return 'Today';
  if (diffDays === 1) return 'Tomorrow';
  if (diffDays <= 6) return date.toLocaleDateString(undefined, { weekday: 'long' });
  if (date.getMonth() === now.getMonth() && date.getFullYear() === now.getFullYear()) {
    return 'Later this month';
  }
  return date.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
}

function groupByRelativeDate(events: IEventItem[]): IGroupedEvents[] {
  const now = new Date();
  const groups = new Map<string, IEventItem[]>();
  const order: string[] = [];

  for (const event of events) {
    const label = getRelativeDateLabel(new Date(event.startDate), now);
    if (!groups.has(label)) {
      groups.set(label, []);
      order.push(label);
    }
    groups.get(label)!.push(event);
  }

  return order.map(label => ({ label, events: groups.get(label)! }));
}

const ListView: React.FC<IListViewProps> = ({ events, availableFields, cardDisplay, onEventClick }) => {
  const grouped = groupByRelativeDate(events);

  if (events.length === 0) {
    return (
      <div className={styles.empty}>
        No upcoming events
      </div>
    );
  }

  return (
    <div className={styles.listView}>
      {grouped.map(group => (
        <div key={group.label} className={styles.group}>
          <div className={styles.dateHeader}>{group.label}</div>
          {group.events.map(event => (
            <EventCard
              key={event.id}
              event={event}
              availableFields={availableFields}
              cardDisplay={cardDisplay}
              layout="compact"
              onClick={onEventClick}
            />
          ))}
        </div>
      ))}
    </div>
  );
};

export default ListView;
