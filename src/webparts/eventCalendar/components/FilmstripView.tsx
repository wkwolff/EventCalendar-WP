/**
 * @file FilmstripView.tsx
 * @description Horizontally scrollable filmstrip layout for event cards. Renders
 *   cards in a single-row track with CSS scroll-snap behavior and dot-style
 *   pagination indicators. Scroll position syncs with the active page dot.
 * @author W. Kevin Wolff
 * @copyright Wolff Creative
 */

import * as React from 'react';
import { IEventItem } from '../models/IEventItem';
import { IFieldInfo } from '../models/IFieldInfo';
import { ICardDisplayOptions } from './IEventCalendarProps';
import EventCard from './EventCard';
import styles from './FilmstripView.module.scss';

/**
 * Props for the FilmstripView component.
 */
export interface IFilmstripViewProps {
  /** Array of event items to render as filmstrip cards. */
  events: IEventItem[];
  /** Metadata for user-selected display fields. */
  availableFields: IFieldInfo[];
  /** Card display toggles from the property pane. */
  cardDisplay: ICardDisplayOptions;
  /** Callback invoked when a filmstrip card is clicked. */
  onEventClick: (event: IEventItem) => void;
}

/** Number of cards visible per "page" in the filmstrip pagination. */
const CARDS_PER_PAGE = 3;

/**
 * Renders event cards in a horizontal scrolling filmstrip with dot pagination.
 * The track uses native horizontal overflow scrolling. Pagination dots are
 * computed from the total card count divided by `CARDS_PER_PAGE`.
 *
 * @param props - Filmstrip view configuration and event data.
 * @returns A horizontally scrollable filmstrip with optional pagination dots.
 */
const FilmstripView: React.FC<IFilmstripViewProps> = ({
  events,
  availableFields,
  cardDisplay,
  onEventClick,
}) => {
  const trackRef = React.useRef<HTMLDivElement>(null);
  const [activePage, setActivePage] = React.useState(0);
  const totalPages = Math.max(1, Math.ceil(events.length / CARDS_PER_PAGE));

  /**
   * Programmatically scrolls the track to a specific page and updates
   * the active pagination dot.
   * @param page - Zero-based page index to scroll to.
   */
  const scrollToPage = (page: number): void => {
    if (!trackRef.current) return;
    // Card width (260px) + gap (16px) = total slot width per card
    const cardWidth = 260 + 16;
    trackRef.current.scrollLeft = page * CARDS_PER_PAGE * cardWidth;
    setActivePage(page);
  };

  /**
   * Handles native scroll events on the track to keep the active pagination
   * dot in sync when the user scrolls manually (e.g., via mouse wheel or touch).
   */
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
      {/* Horizontally scrollable card track */}
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

      {/* Dot pagination — only shown when there are multiple pages */}
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
