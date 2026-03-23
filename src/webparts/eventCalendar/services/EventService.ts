/**
 * @file EventService.ts
 * @description Service layer responsible for fetching calendar events from a
 *              SharePoint list via PnPjs and normalizing them into
 *              {@link IEventItem} objects. Handles dynamic column mapping,
 *              image extraction from various SharePoint column types, and
 *              field-level safety filtering.
 * @author W. Kevin Wolff
 * @copyright Wolff Creative
 */

import { getSP } from './PnPSetup';
import { IEventItem } from '../models/IEventItem';

/**
 * Fields that require `$expand` or are otherwise incompatible with a plain
 * `$select` projection in the SP REST API. Including them in the query
 * causes a 400 response, so they are silently stripped before the request.
 */
const BLOCKED_SELECT_FIELDS = new Set([
  'ParticipantsPicker',
  'ParticipantsPickerId',
]);

/**
 * Describes which SharePoint list columns map to the core event properties.
 * Values are **internal names** of the list columns.
 */
export interface IEventFieldMapping {
  /** Column whose value becomes the event title. */
  titleField: string;
  /** DateTime column used as the event start. */
  startDateField: string;
  /** DateTime column used as the event end (may be empty). */
  endDateField: string;
  /** Boolean column indicating an all-day event (may be empty). */
  allDayField: string;
  /** Choice/Text column used for the category label (may be empty). */
  categoryField: string;
  /** Text column used for the event location (may be empty). */
  locationField: string;
}

/** Regex for common image file extensions, including optional query strings. */
const IMAGE_EXT_REGEX = /\.(jpg|jpeg|png|gif|bmp|webp|svg)(\?|$)/i;

/**
 * Tests whether a URL string points to a recognized image file.
 *
 * @param url - The URL to test.
 * @returns `true` if the URL ends with a known image extension.
 */
function isImageUrl(url: string): boolean {
  return IMAGE_EXT_REGEX.test(url);
}

/**
 * Attempts to extract an image URL from an arbitrary SharePoint column value.
 *
 * SharePoint stores images differently depending on column type:
 * - **Image column**: JSON string with `serverRelativeUrl` or `Url` property.
 * - **Hyperlink column**: Plain URL string.
 * - **Rich text / other**: May contain an object with URL properties.
 *
 * @param value - The raw column value from the REST response.
 * @returns A resolved image URL, or an empty string if no image is found.
 */
function extractImageUrl(value: unknown): string {
  if (!value) return '';

  if (typeof value === 'string') {
    // Attempt to parse as JSON (Image column stores a JSON blob)
    try {
      const parsed = JSON.parse(value);
      const url = (parsed && (parsed.serverRelativeUrl || parsed.Url)) || '';
      if (url && isImageUrl(url)) return url;
    } catch {
      // Not JSON — treat as a plain URL string (Hyperlink column)
      if (isImageUrl(value)) return value;
    }
  }

  // Handle object values (e.g., already-parsed Image column metadata)
  if (typeof value === 'object' && value !== null) {
    const obj = value as Record<string, unknown>;
    const url = (obj.serverRelativeUrl as string) || (obj.Url as string) || '';
    if (url && isImageUrl(url)) return url;
  }

  return '';
}

/**
 * Fetches events from the specified SharePoint list and normalizes them into
 * an array of {@link IEventItem} objects.
 *
 * The function:
 * 1. Strips blocked fields that cannot be used in `$select`.
 * 2. Builds a deduplicated `$select` clause from core mapped fields plus
 *    user-selected extra fields.
 * 3. Queries the list ordered by start date (ascending), limited to `maxEvents`.
 * 4. Maps raw items to {@link IEventItem}, extracting an image URL from the
 *    first extra field that contains one.
 *
 * @param listId        - GUID of the target SharePoint list.
 * @param fieldMapping  - Column mapping configuration for core event properties.
 * @param selectedFields - Additional column internal names the user wants displayed.
 * @param maxEvents     - Maximum number of items to retrieve.
 * @returns A promise resolving to the normalized event array.
 *
 * @example
 * ```ts
 * const events = await fetchEvents(
 *   'b1c2d3e4-...',
 *   { titleField: 'Title', startDateField: 'EventDate', endDateField: 'EndDate',
 *     allDayField: 'fAllDayEvent', categoryField: 'Category', locationField: 'Location' },
 *   ['Description', 'BannerImage'],
 *   100
 * );
 * ```
 */
export async function fetchEvents(
  listId: string,
  fieldMapping: IEventFieldMapping,
  selectedFields: string[],
  maxEvents: number
): Promise<IEventItem[]> {
  const sp = getSP();

  // Remove fields that are known to break REST $select queries
  const safeFields = selectedFields.filter(f => !BLOCKED_SELECT_FIELDS.has(f));

  // Build the core $select list — these fields are always needed for rendering
  const coreSelect = ['Id', fieldMapping.titleField, fieldMapping.startDateField];
  if (fieldMapping.endDateField) coreSelect.push(fieldMapping.endDateField);
  if (fieldMapping.allDayField) coreSelect.push(fieldMapping.allDayField);
  if (fieldMapping.categoryField) coreSelect.push(fieldMapping.categoryField);
  if (fieldMapping.locationField) coreSelect.push(fieldMapping.locationField);

  // Deduplicate: user-selected fields that already appear in core are skipped
  const coreSet = new Set(coreSelect);
  const extraFields = safeFields.filter(f => !coreSet.has(f));
  const selectFields = [...coreSelect, ...extraFields];

  // Execute the REST query with $select, $orderby, and $top
  const items = await sp.web.lists.getById(listId).items
    .select(...selectFields)
    .orderBy(fieldMapping.startDateField, true)
    .top(maxEvents)();

  // Map each raw SP item into the normalized IEventItem shape
  return items.map((item: Record<string, unknown>) => {
    // Collect extra field values into a generic bag for the detail panel
    const fields: Record<string, unknown> = {};
    for (const f of extraFields) {
      fields[f] = item[f];
    }

    // Scan extra fields for the first value that looks like an image URL
    let imageUrl = '';
    for (const f of extraFields) {
      const url = extractImageUrl(item[f]);
      if (url) {
        imageUrl = url;
        break;
      }
    }

    return {
      id: item.Id as number,
      title: (item[fieldMapping.titleField] as string) || '',
      startDate: item[fieldMapping.startDateField] as string,
      // Fall back to startDate when no end-date column is mapped or the value is null
      endDate: fieldMapping.endDateField
        ? (item[fieldMapping.endDateField] as string) || (item[fieldMapping.startDateField] as string)
        : item[fieldMapping.startDateField] as string,
      allDay: fieldMapping.allDayField ? !!item[fieldMapping.allDayField] : false,
      category: fieldMapping.categoryField ? (item[fieldMapping.categoryField] as string) || '' : '',
      location: fieldMapping.locationField ? (item[fieldMapping.locationField] as string) || '' : '',
      imageUrl,
      fields,
    };
  });
}
