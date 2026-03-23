/**
 * @file ListView.tsx
 * @description Compact list view that groups events by relative date labels
 *   ("Today", "Tomorrow", weekday names, "Later this month", or "Month Year").
 *   Each group renders a date header followed by compact EventCard rows.
 * @author W. Kevin Wolff
 * @copyright TidalHealth
 */

import * as React from 'react';
import { IEventItem } from '../models/IEventItem';
import { IFieldInfo } from '../models/IFieldInfo';
import { ICardDisplayOptions } from './IEventCalendarProps';
import EventCard from './EventCard';
import styles from './ListView.module.scss';

/**
 * Props for the ListView component.
 */
export interface IListViewProps {
  /** Array of event items to render, pre-sorted by start date. */
  events: IEventItem[];
  /** Metadata for user-selected display fields. */
  availableFields: IFieldInfo[];
  /** Card display toggles from the property pane. */
  cardDisplay: ICardDisplayOptions;
  /** Callback invoked when a compact card is clicked. */
  onEventClick: (event: IEventItem) => void;
}

/**
 * Internal interface representing a group of events under a shared date label.
 */
interface IGroupedEvents {
  /** Human-readable date label (e.g., "Today", "Tomorrow", "Later this month"). */
  label: string;
  /** Events belonging to this group. */
  events: IEventItem[];
}

/**
 * Computes a human-friendly relative date label for an event date.
 * The progression is: "Earlier" -> "Today" -> "Tomorrow" -> weekday name
 * (within 6 days) -> "Later this month" -> "Month Year" (future months).
 *
 * @param date - The event's start date.
 * @param now - The current date/time reference.
 * @returns A relative date label string.
 */
function getRelativeDateLabel(date: Date, now: Date): string {
  // Normalize both dates to midnight for clean day-difference calculation
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

/**
 * Groups a sorted array of events by their relative date label, preserving
 * insertion order so groups appear chronologically.
 *
 * @param events - Pre-sorted event items.
 * @returns An ordered array of grouped events.
 */
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

  // Reconstruct as an ordered array using the insertion-order tracking array
  return order.map(label => ({ label, events: groups.get(label)! }));
}

/**
 * Renders events in a vertically stacked compact list, grouped under relative
 * date section headers. Shows an empty-state message when no events exist.
 *
 * @param props - List view configuration and event data.
 * @returns A grouped list of compact event cards, or an empty-state message.
 */
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
