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
import { IFieldInfo } from '../models/IFieldInfo';

/** Field types that require $expand to retrieve their sub-properties (e.g., Title, EMail). */
const EXPANDABLE_FIELD_TYPES = new Set(['User', 'UserMulti', 'Lookup', 'LookupMulti']);

/**
 * Fields that require `$expand` or are otherwise incompatible with a plain
 * `$select` projection in the SP REST API. Including them in the query
 * causes a 400 response, so they are silently stripped before the request.
 */
const BLOCKED_SELECT_FIELDS = new Set([
  'ParticipantsPicker',
  'ParticipantsPickerId',
  'ItemChildCount',
  'FolderChildCount',
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
  maxEvents: number,
  availableFields?: IFieldInfo[]
): Promise<IEventItem[]> {
  const sp = getSP();

  // Build a type lookup from available field metadata
  const fieldTypeMap = new Map<string, string>();
  if (availableFields) {
    for (const f of availableFields) {
      fieldTypeMap.set(f.internalName, f.fieldType);
    }
  }

  // Remove fields that are known to break REST $select queries
  const safeFields = selectedFields.filter(f => !BLOCKED_SELECT_FIELDS.has(f));

  // Build the core $select list — these fields are always needed for rendering
  // Include 'Attachments' boolean so we know which items have files to fetch
  const coreSelect = ['Id', 'Attachments', fieldMapping.titleField, fieldMapping.startDateField];
  if (fieldMapping.endDateField) coreSelect.push(fieldMapping.endDateField);
  if (fieldMapping.allDayField) coreSelect.push(fieldMapping.allDayField);
  if (fieldMapping.categoryField) coreSelect.push(fieldMapping.categoryField);
  if (fieldMapping.locationField) coreSelect.push(fieldMapping.locationField);

  // Deduplicate: user-selected fields that already appear in core are skipped
  const coreSet = new Set(coreSelect);
  const extraFields = safeFields.filter(f => !coreSet.has(f));

  // Separate expandable fields (User, Lookup) from plain fields
  const expandFields: string[] = [];
  const plainFields: string[] = [];
  for (const f of extraFields) {
    const fType = fieldTypeMap.get(f) || '';
    if (EXPANDABLE_FIELD_TYPES.has(fType)) {
      expandFields.push(f);
    } else {
      plainFields.push(f);
    }
  }

  // Build $select: plain fields use their name directly;
  // expandable fields use FieldName/Title (and /EMail for User types)
  const selectFields = [...coreSelect, ...plainFields];
  for (const f of expandFields) {
    const fType = fieldTypeMap.get(f) || '';
    selectFields.push(f + '/Title');
    if (fType === 'User' || fType === 'UserMulti') {
      selectFields.push(f + '/EMail');
    }
  }

  // Build the query — add $expand for lookup/person fields
  let query = sp.web.lists.getById(listId).items
    .select(...selectFields)
    .orderBy(fieldMapping.startDateField, true)
    .top(maxEvents);

  if (expandFields.length > 0) {
    query = query.expand(...expandFields);
  }

  const items = await query();

  // Map each raw SP item into the normalized IEventItem shape
  const mapped = items.map((item: Record<string, unknown>) => {
    // Collect extra field values into a generic bag for the detail panel.
    // Expanded fields (User/Lookup) return objects like { Title, EMail } —
    // flatten them into a display-friendly string.
    const fields: Record<string, unknown> = {};
    for (const f of plainFields) {
      fields[f] = item[f];
    }
    for (const f of expandFields) {
      const expanded = item[f] as Record<string, unknown> | null;
      if (expanded && expanded.Title) {
        fields[f] = expanded.Title as string;
      } else {
        fields[f] = null;
      }
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
      hasAttachments: !!item.Attachments,
      attachments: [] as Array<{ fileName: string; url: string }>,
    };
  });

  // Fetch attachment details for items that have them.
  // Sequential loop to satisfy require-atomic-updates lint rule.
  const itemsWithAttachments = mapped.filter(e => e.hasAttachments);
  for (const evt of itemsWithAttachments) {
    try {
      const atts = await sp.web.lists.getById(listId).items
        .getById(evt.id).attachmentFiles() as Array<{ FileName: string; ServerRelativeUrl: string }>;
      evt.attachments = atts.map(a => ({
        fileName: a.FileName,
        url: a.ServerRelativeUrl,
      }));
      // Use first image attachment as card image if none was found from fields
      if (!evt.imageUrl) {
        const imgAtt = atts.find(a => isImageUrl(a.FileName));
        if (imgAtt) evt.imageUrl = imgAtt.ServerRelativeUrl;
      }
    } catch {
      // Silently skip — attachments are supplementary
    }
  }

  // Strip the temporary hasAttachments flag
  return mapped.map(({ hasAttachments: _, ...rest }) => rest);
}
