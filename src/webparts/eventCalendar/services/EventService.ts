import { getSP } from './PnPSetup';
import { IEventItem } from '../models/IEventItem';

// Fields that require $expand or are otherwise incompatible with plain $select
const BLOCKED_SELECT_FIELDS = new Set([
  'ParticipantsPicker',
  'ParticipantsPickerId',
]);

export interface IEventFieldMapping {
  titleField: string;
  startDateField: string;
  endDateField: string;
  allDayField: string;
  categoryField: string;
  locationField: string;
}

function extractImageUrl(value: unknown): string {
  if (!value) return '';
  if (typeof value === 'string') {
    // Could be a JSON string from an Image column
    try {
      const parsed = JSON.parse(value);
      if (parsed && parsed.serverRelativeUrl) return parsed.serverRelativeUrl;
      if (parsed && parsed.Url) return parsed.Url;
    } catch {
      // Plain URL string
      if (/\.(jpg|jpeg|png|gif|bmp|webp|svg)(\?|$)/i.test(value)) return value;
    }
  }
  if (typeof value === 'object' && value !== null) {
    const obj = value as Record<string, unknown>;
    if (obj.serverRelativeUrl) return obj.serverRelativeUrl as string;
    if (obj.Url) return obj.Url as string;
  }
  return '';
}

export async function fetchEvents(
  listId: string,
  fieldMapping: IEventFieldMapping,
  selectedFields: string[],
  maxEvents: number
): Promise<IEventItem[]> {
  const sp = getSP();
  const safeFields = selectedFields.filter(f => !BLOCKED_SELECT_FIELDS.has(f));

  // Build core select from mapped fields
  const coreSelect = ['Id', fieldMapping.titleField, fieldMapping.startDateField];
  if (fieldMapping.endDateField) coreSelect.push(fieldMapping.endDateField);
  if (fieldMapping.allDayField) coreSelect.push(fieldMapping.allDayField);
  if (fieldMapping.categoryField) coreSelect.push(fieldMapping.categoryField);
  if (fieldMapping.locationField) coreSelect.push(fieldMapping.locationField);

  // Deduplicate: remove any display fields that overlap with core
  const coreSet = new Set(coreSelect);
  const extraFields = safeFields.filter(f => !coreSet.has(f));
  const selectFields = [...coreSelect, ...extraFields];

  const items = await sp.web.lists.getById(listId).items
    .select(...selectFields)
    .orderBy(fieldMapping.startDateField, true)
    .top(maxEvents)();

  return items.map((item: Record<string, unknown>) => {
    const fields: Record<string, unknown> = {};
    for (const f of extraFields) {
      fields[f] = item[f];
    }

    // Try to find an image from the extra fields
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
