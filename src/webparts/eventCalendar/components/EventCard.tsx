/**
 * @file EventCard.tsx
 * @description Renders individual event cards in either "filmstrip" (vertical card with
 *   hero image) or "compact" (horizontal row with date block and optional thumbnail)
 *   layout. Handles date formatting, image resolution from extra fields, and renders
 *   additional user-selected fields via FieldBadge components.
 * @author W. Kevin Wolff
 * @copyright TidalHealth
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IEventItem } from '../models/IEventItem';
import { IFieldInfo } from '../models/IFieldInfo';
import { ICardDisplayOptions } from './IEventCalendarProps';
import FieldBadge, { isEmptyValue, isImageFieldValue, getImageUrl } from './FieldBadge';
import styles from './EventCard.module.scss';

/**
 * Props for the EventCard component.
 */
export interface IEventCardProps {
  /** The event data to render. */
  event: IEventItem;
  /** Metadata for user-selected display fields (non-core). */
  availableFields: IFieldInfo[];
  /** Card display toggles from the property pane. */
  cardDisplay: ICardDisplayOptions;
  /** Which card layout variant to render. */
  layout: 'filmstrip' | 'compact';
  /** Callback invoked when the card is clicked or activated via keyboard. */
  onClick: (event: IEventItem) => void;
}

/**
 * Formats a start/end date pair into a human-readable time range string.
 * Uses the browser locale for time formatting (e.g., "2:00 PM - 4:30 PM").
 * @param start - ISO date string for the event start.
 * @param end - ISO date string for the event end.
 * @returns A formatted time range like "2:00 PM - 4:30 PM".
 */
function formatTimeRange(start: string, end: string): string {
  const s = new Date(start);
  const e = new Date(end);
  const startTime = s.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  const endTime = e.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  return startTime + ' - ' + endTime;
}

/**
 * Formats a start/end date pair into a day-of-week plus time range string.
 * @param start - ISO date string for the event start.
 * @param end - ISO date string for the event end.
 * @returns A string like "Wednesday 2:00 PM - 4:30 PM".
 */
function formatDayTime(start: string, end: string): string {
  const s = new Date(start);
  const dayName = s.toLocaleDateString(undefined, { weekday: 'long' });
  return dayName + ' ' + formatTimeRange(start, end);
}

/**
 * Returns the full month name for a given date (e.g., "January").
 * @param date - The date to extract the month from.
 * @returns Full month name string.
 */
function getMonthName(date: Date): string {
  return date.toLocaleDateString(undefined, { month: 'long' });
}

/**
 * Returns the abbreviated, uppercased month name for a given date (e.g., "JAN").
 * Used in the compact card's date block.
 * @param date - The date to extract the month abbreviation from.
 * @returns Uppercased abbreviated month string.
 */
function getMonthAbbr(date: Date): string {
  return date.toLocaleDateString(undefined, { month: 'short' }).toUpperCase();
}

/**
 * Filmstrip card variant — a vertical card with an optional hero image at the top,
 * followed by month/day, category badge, title, time range, location, and extra fields.
 *
 * @param props - Event card props.
 * @returns A filmstrip-style event card element.
 */
const FilmstripCard: React.FC<IEventCardProps> = ({ event, availableFields, cardDisplay, onClick }) => {
  const startDate = new Date(event.startDate);

  // Build a lookup map for O(1) field metadata access by internal name
  const fieldMap = new Map(availableFields.map(f => [f.internalName, f]));

  // Collect non-empty, non-image extra fields for display at the bottom of the card.
  // Category and location have dedicated rendering spots, so image fields are excluded
  // here to avoid duplication — they are handled by the hero image slot.
  const visibleFields: Array<{ key: string; field: IFieldInfo; value: unknown }> = [];
  Object.keys(event.fields).forEach((key: string) => {
    const value = event.fields[key];
    const field = fieldMap.get(key);
    if (!field || isEmptyValue(value, field.fieldType)) return;
    if (isImageFieldValue(field, value)) return;
    visibleFields.push({ key, field, value });
  });

  return (
    <div
      className={styles.filmstripCard}
      onClick={() => onClick(event)}
      role="button"
      tabIndex={0}
      onKeyDown={(e) => { if (e.key === 'Enter') onClick(event); }}
    >
      {/* Hero image area — shows the event banner or a calendar icon placeholder */}
      {cardDisplay.showImage && (
        <div className={styles.cardImage}>
          {event.imageUrl ? (
            <img src={event.imageUrl} alt={event.title} />
          ) : (
            <Icon iconName="Calendar" className={styles.cardImagePlaceholder} />
          )}
        </div>
      )}

      <div className={styles.cardBody}>
        <div className={styles.cardMonth}>{getMonthName(startDate)}</div>
        {/* Zero-pad single-digit days for visual consistency */}
        <div className={styles.cardDay}>{startDate.getDate() < 10 ? '0' + startDate.getDate() : String(startDate.getDate())}</div>

        {cardDisplay.showCategory && event.category && (
          <div className={styles.cardCategory}>{event.category}</div>
        )}

        <div className={styles.cardTitle} title={event.title}>{event.title}</div>

        {cardDisplay.showTime && (
          <div className={styles.cardTime}>
            {event.allDay ? 'All day' : formatDayTime(event.startDate, event.endDate)}
          </div>
        )}

        {cardDisplay.showLocation && event.location && (
          <div className={styles.cardLocation}>{event.location}</div>
        )}

        {/* Extra user-selected fields rendered as compact badges */}
        {visibleFields.length > 0 && (
          <div className={styles.fields}>
            {visibleFields.map((item) => (
              <FieldBadge key={item.key} field={item.field} value={item.value} />
            ))}
          </div>
        )}
      </div>
    </div>
  );
};

/**
 * Compact card variant — a horizontal row with a date block on the left,
 * event metadata in the center, and an optional thumbnail on the right.
 * Image fields from extra data are promoted to the thumbnail slot if no
 * primary `imageUrl` is available.
 *
 * @param props - Event card props.
 * @returns A compact-style event card element.
 */
const CompactCard: React.FC<IEventCardProps> = ({ event, availableFields, cardDisplay, onClick }) => {
  const startDate = new Date(event.startDate);
  const fieldMap = new Map(availableFields.map(f => [f.internalName, f]));

  // Start with the event's primary image; fall back to the first image-type extra field
  let bannerUrl = cardDisplay.showImage ? event.imageUrl : '';
  const visibleFields: Array<{ key: string; field: IFieldInfo; value: unknown }> = [];

  Object.keys(event.fields).forEach((key: string) => {
    const value = event.fields[key];
    const field = fieldMap.get(key);
    if (!field || isEmptyValue(value, field.fieldType)) return;
    if (isImageFieldValue(field, value)) {
      // Promote the first image-type field to the thumbnail when no primary image exists
      if (!bannerUrl) {
        const url = getImageUrl(value);
        if (url) { bannerUrl = url; return; }
      }
      return;
    }
    visibleFields.push({ key, field, value });
  });

  return (
    <div
      className={styles.compactCard}
      onClick={() => onClick(event)}
      role="button"
      tabIndex={0}
      onKeyDown={(e) => { if (e.key === 'Enter') onClick(event); }}
    >
      {/* Left-side date block showing abbreviated month and day number */}
      <div className={styles.dateBlock}>
        <span className={styles.dateMonth}>{getMonthAbbr(startDate)}</span>
        <span className={styles.dateDay}>{startDate.getDate()}</span>
      </div>

      <div className={styles.compactContent}>
        {cardDisplay.showCategory && event.category && (
          <div className={styles.compactCategory}>{event.category}</div>
        )}
        <div className={styles.compactTitle}>{event.title}</div>
        {cardDisplay.showTime && (
          <div className={styles.compactTime}>
            {event.allDay ? 'All day' : formatDayTime(event.startDate, event.endDate)}
          </div>
        )}
        {cardDisplay.showLocation && event.location && (
          <div className={styles.compactLocation}>
            <Icon iconName="POI" className={styles.locationIcon} />
            {event.location}
          </div>
        )}
        {visibleFields.length > 0 && (
          <div className={styles.fields}>
            {visibleFields.map((item) => (
              <FieldBadge key={item.key} field={item.field} value={item.value} />
            ))}
          </div>
        )}
      </div>

      {/* Right-side thumbnail — sourced from primary image or promoted extra field */}
      {bannerUrl && (
        <div className={styles.compactThumbnail}>
          <img src={bannerUrl} alt={event.title} />
        </div>
      )}
    </div>
  );
};

/**
 * Composite EventCard component that delegates to the appropriate layout variant
 * based on the `layout` prop.
 *
 * @param props - Event card props including the layout discriminator.
 * @returns Either a FilmstripCard or CompactCard element.
 */
const EventCard: React.FC<IEventCardProps> = (props) => {
  if (props.layout === 'filmstrip') {
    return <FilmstripCard {...props} />;
  }
  return <CompactCard {...props} />;
};

export default EventCard;
