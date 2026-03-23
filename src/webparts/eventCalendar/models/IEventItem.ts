/**
 * @file IEventItem.ts
 * @description Normalized event item model used throughout the calendar web part.
 *              Represents a single calendar event after it has been fetched and
 *              mapped from a SharePoint list item into a framework-agnostic shape.
 * @author W. Kevin Wolff
 * @copyright Wolff Creative
 */

/**
 * A normalized calendar event derived from a SharePoint list item.
 *
 * Core scheduling properties (`startDate`, `endDate`, `allDay`) drive the
 * FullCalendar rendering, while `fields` carries any additional columns the
 * user chose to display in the detail panel.
 *
 * @example
 * ```ts
 * const evt: IEventItem = {
 *   id: 42,
 *   title: 'Board Meeting',
 *   startDate: '2026-03-23T09:00:00Z',
 *   endDate: '2026-03-23T10:00:00Z',
 *   allDay: false,
 *   category: 'Meeting',
 *   location: 'Room 204',
 *   imageUrl: '',
 *   fields: { Description: 'Quarterly review' },
 * };
 * ```
 */
export interface IEventItem {
  /** SharePoint list item ID. */
  id: number;

  /** Event display title sourced from the mapped title column. */
  title: string;

  /** ISO 8601 start date/time string. */
  startDate: string;

  /** ISO 8601 end date/time string; falls back to `startDate` when no end date column is mapped. */
  endDate: string;

  /** Whether the event spans entire day(s), hiding time-of-day display. */
  allDay: boolean;

  /** Optional category or type label used for color-coding and filtering. */
  category: string;

  /** Optional location or venue text. */
  location: string;

  /** URL to a thumbnail image extracted from an Image/Hyperlink column, or empty string. */
  imageUrl: string;

  /** Bag of additional column values selected by the user for the detail view. */
  fields: Record<string, unknown>;

  /** File attachments on the list item (fetched separately from item fields). */
  attachments: Array<{ fileName: string; url: string }>;
}
