import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IEventItem } from '../models/IEventItem';
import { IFieldInfo } from '../models/IFieldInfo';
import { ICardDisplayOptions } from './IEventCalendarProps';
import FieldBadge, { isEmptyValue, isImageFieldValue, getImageUrl } from './FieldBadge';
import styles from './EventCard.module.scss';

export interface IEventCardProps {
  event: IEventItem;
  availableFields: IFieldInfo[];
  cardDisplay: ICardDisplayOptions;
  layout: 'filmstrip' | 'compact';
  onClick: (event: IEventItem) => void;
}

function formatTimeRange(start: string, end: string): string {
  const s = new Date(start);
  const e = new Date(end);
  const startTime = s.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  const endTime = e.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  return startTime + ' - ' + endTime;
}

function formatDayTime(start: string, end: string): string {
  const s = new Date(start);
  const dayName = s.toLocaleDateString(undefined, { weekday: 'long' });
  return dayName + ' ' + formatTimeRange(start, end);
}

function getMonthName(date: Date): string {
  return date.toLocaleDateString(undefined, { month: 'long' });
}

function getMonthAbbr(date: Date): string {
  return date.toLocaleDateString(undefined, { month: 'short' }).toUpperCase();
}

const FilmstripCard: React.FC<IEventCardProps> = ({ event, availableFields, cardDisplay, onClick }) => {
  const startDate = new Date(event.startDate);
  const fieldMap = new Map(availableFields.map(f => [f.internalName, f]));

  // Gather extra visible fields (excluding category/location which have dedicated spots)
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

const CompactCard: React.FC<IEventCardProps> = ({ event, availableFields, cardDisplay, onClick }) => {
  const startDate = new Date(event.startDate);
  const fieldMap = new Map(availableFields.map(f => [f.internalName, f]));

  let bannerUrl = cardDisplay.showImage ? event.imageUrl : '';
  const visibleFields: Array<{ key: string; field: IFieldInfo; value: unknown }> = [];

  Object.keys(event.fields).forEach((key: string) => {
    const value = event.fields[key];
    const field = fieldMap.get(key);
    if (!field || isEmptyValue(value, field.fieldType)) return;
    if (isImageFieldValue(field, value)) {
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

      {bannerUrl && (
        <div className={styles.compactThumbnail}>
          <img src={bannerUrl} alt={event.title} />
        </div>
      )}
    </div>
  );
};

const EventCard: React.FC<IEventCardProps> = (props) => {
  if (props.layout === 'filmstrip') {
    return <FilmstripCard {...props} />;
  }
  return <CompactCard {...props} />;
};

export default EventCard;
