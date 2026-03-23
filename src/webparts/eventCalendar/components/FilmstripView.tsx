import * as React from 'react';
import { IEventItem } from '../models/IEventItem';
import { IFieldInfo } from '../models/IFieldInfo';
import { ICardDisplayOptions } from './IEventCalendarProps';
import EventCard from './EventCard';
import styles from './FilmstripView.module.scss';

export interface IFilmstripViewProps {
  events: IEventItem[];
  availableFields: IFieldInfo[];
  cardDisplay: ICardDisplayOptions;
  onEventClick: (event: IEventItem) => void;
}

const CARDS_PER_PAGE = 3;

const FilmstripView: React.FC<IFilmstripViewProps> = ({
  events,
  availableFields,
  cardDisplay,
  onEventClick,
}) => {
  const trackRef = React.useRef<HTMLDivElement>(null);
  const [activePage, setActivePage] = React.useState(0);
  const totalPages = Math.max(1, Math.ceil(events.length / CARDS_PER_PAGE));

  const scrollToPage = (page: number): void => {
    if (!trackRef.current) return;
    const cardWidth = 260 + 16; // card width + gap
    trackRef.current.scrollLeft = page * CARDS_PER_PAGE * cardWidth;
    setActivePage(page);
  };

  const handleScroll = (): void => {
    if (!trackRef.current) return;
    const cardWidth = 260 + 16;
    const page = Math.round(trackRef.current.scrollLeft / (CARDS_PER_PAGE * cardWidth));
    setActivePage(Math.min(page, totalPages - 1));
  };

  if (events.length === 0) {
    return <div className={styles.empty}>No upcoming events</div>;
  }

  return (
    <div className={styles.filmstripView}>
      <div
        className={styles.track}
        ref={trackRef}
        onScroll={handleScroll}
      >
        {events.map(event => (
          <EventCard
            key={event.id}
            event={event}
            availableFields={availableFields}
            cardDisplay={cardDisplay}
            layout="filmstrip"
            onClick={onEventClick}
          />
        ))}
      </div>

      {totalPages > 1 && (
        <div className={styles.pagination}>
          {Array.from({ length: totalPages }, (_, i) => (
            <button
              key={i}
              className={`${styles.dot} ${i === activePage ? styles.dotActive : ''}`}
              onClick={() => scrollToPage(i)}
              aria-label={`Page ${i + 1}`}
            />
          ))}
        </div>
      )}
    </div>
  );
};

export default FilmstripView;
