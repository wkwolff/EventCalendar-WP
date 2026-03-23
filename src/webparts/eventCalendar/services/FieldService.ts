/**
 * @file FieldService.ts
 * @description Service layer for retrieving and analyzing SharePoint list field
 *              (column) metadata. Provides field fetching, display-field
 *              filtering, and intelligent auto-detection of standard event
 *              column mappings so the property pane can offer sensible defaults.
 * @author W. Kevin Wolff
 * @copyright Wolff Creative
 */

import { getSP } from './PnPSetup';
import { IFieldInfo } from '../models/IFieldInfo';

/**
 * Internal/system field names that should never appear in the property-pane
 * dropdowns. These are SharePoint infrastructure columns that have no
 * meaningful user-facing value for a calendar web part.
 */
const SYSTEM_FIELDS = new Set([
  'ContentType', 'Attachments', '_ModerationComments', 'File_x0020_Type',
  'MetaInfo', 'DocIcon', '_HasCopyDestinations', '_CopySource',
  'owshiddenversion', 'WorkflowVersion', '_UIVersion', '_UIVersionString',
  'InstanceID', 'Order', 'GUID', 'WorkflowInstanceID', 'FileRef',
  'FileDirRef', 'Last_x0020_Modified', 'Created_x0020_Date',
  'FSObjType', 'SortBehavior', 'PermMask', 'FileLeafRef',
  'UniqueId', 'SyncClientId', 'ProgId', 'ScopeId', 'HTML_x0020_File_x0020_Type',
  '_EditMenuTableStart', '_EditMenuTableStart2', '_EditMenuTableEnd',
  'LinkFilenameNoMenu', 'LinkFilename', 'LinkFilename2',
  'SelectTitle', 'Edit', 'AppAuthor', 'AppEditor',
  'ComplianceAssetId', '_ComplianceFlags', '_ComplianceTag',
  '_ComplianceTagWrittenTime', '_ComplianceTagUserId',
  '_IsRecord', 'AccessPolicy', '_VirusStatus', '_VirusVendorID',
  '_VirusInfo', '_CommentFlags', '_CommentCount',
]);

/** Field types that cannot be projected via REST `$select`. */
const UNSUPPORTED_FIELD_TYPES = new Set([
  'Geolocation',
]);

/** Specific field names known to cause REST API errors. */
const UNSUPPORTED_FIELD_NAMES = new Set([
  'ParticipantsPicker',
  'ParticipantsPickerId',
]);

/**
 * Fetches lists that are likely event/calendar lists by checking for the
 * presence of DateTime fields. Includes both classic Calendar lists (template 106)
 * and modern Events lists (template 100 with date columns).
 *
 * @returns A promise resolving to an array of `{ id, title }` for eligible lists.
 */
export async function fetchEventLists(): Promise<Array<{ id: string; title: string }>> {
  const sp = getSP();

  // Fetch all non-hidden lists with their base template
  const lists = await sp.web.lists
    .filter("Hidden eq false")
    .select('Id', 'Title', 'BaseTemplate')() as Array<{
      Id: string;
      Title: string;
      BaseTemplate: number;
    }>;

  // Classic Calendar lists (template 106) are always included
  const calendarLists = lists.filter(l => l.BaseTemplate === 106);
  const calendarIds = new Set(calendarLists.map(l => l.Id));

  // For remaining lists, check if they have DateTime fields (event-like lists)
  const candidates = lists.filter(l =>
    l.BaseTemplate !== 106 &&
    !l.Title.startsWith('_') &&
    // Exclude document libraries (101), form libraries, and system lists
    l.BaseTemplate !== 101 &&
    l.BaseTemplate !== 115 &&
    l.BaseTemplate !== 119 &&
    l.BaseTemplate !== 851
  );
  const eventLists = [...calendarLists];

  for (const list of candidates) {
    try {
      const fields = await sp.web.lists.getById(list.Id).fields
        .filter("Hidden eq false and (TypeAsString eq 'DateTime')")
        .select('InternalName')
        .top(1)();
      if (fields.length > 0 && !calendarIds.has(list.Id)) {
        eventLists.push(list);
      }
    } catch {
      // Skip lists we can't query
    }
  }

  return eventLists
    .sort((a, b) => a.Title.localeCompare(b.Title))
    .map(l => ({ id: l.Id, title: l.Title }));
}

/**
 * Fetches all user-facing fields for a SharePoint list, filtering out hidden
 * columns, system infrastructure columns, and field types incompatible with
 * REST queries.
 *
 * Used to populate the column-mapping dropdowns and the additional-fields
 * picker in the web part property pane.
 *
 * @param listId - GUID of the SharePoint list.
 * @returns A promise resolving to an array of {@link IFieldInfo} objects.
 */
export async function fetchAllListFields(listId: string): Promise<IFieldInfo[]> {
  const sp = getSP();

  // Request only non-hidden fields with the minimum projection needed
  const fields = await sp.web.lists.getById(listId).fields
    .filter("Hidden eq false")
    .select('InternalName', 'Title', 'TypeAsString', 'Required')();

  // Apply three layers of filtering: system names, unsupported types, blocked names
  return fields
    .filter((f: { InternalName: string; TypeAsString: string }) =>
      !SYSTEM_FIELDS.has(f.InternalName) &&
      !UNSUPPORTED_FIELD_TYPES.has(f.TypeAsString) &&
      !UNSUPPORTED_FIELD_NAMES.has(f.InternalName)
    )
    .map((f: { InternalName: string; Title: string; TypeAsString: string; Required: boolean }) => ({
      internalName: f.InternalName,
      displayName: f.Title,
      fieldType: f.TypeAsString,
      required: f.Required,
    }));
}

/**
 * Returns the subset of fields eligible for the "additional display fields"
 * picker by excluding columns already assigned to a core mapping slot and
 * the synthetic `ID` column.
 *
 * @param allFields    - Full field list from {@link fetchAllListFields}.
 * @param mappedFields - Internal names of columns currently assigned to core mappings.
 * @returns Filtered array of {@link IFieldInfo} suitable for the display-fields picker.
 */
export function getDisplayFields(
  allFields: IFieldInfo[],
  mappedFields: string[]
): IFieldInfo[] {
  const mapped = new Set(mappedFields);
  return allFields.filter(f =>
    !mapped.has(f.internalName) &&
    f.internalName !== 'ID'
  );
}

/**
 * Inspects the available fields of a list and returns best-guess column
 * mappings for an event calendar.
 *
 * Detection strategy (in priority order):
 * 1. **Standard Events list**: looks for `EventDate` / `EndDate` / `fAllDayEvent`
 *    (the OOB SharePoint "Events" or "Calendar" list schema).
 * 2. **Custom list with recognizable names**: scans DateTime fields for common
 *    names like "StartDate", "Start", "End", "EndDate", etc.
 * 3. **Fallback**: uses the first two DateTime fields found as start/end.
 *
 * Category and location are detected independently by scanning Choice/Text
 * fields for keyword matches.
 *
 * @param fields - The full field list retrieved by {@link fetchAllListFields}.
 * @returns An object with suggested internal-name mappings for each core slot.
 *
 * @example
 * ```ts
 * const mappings = autoDetectFieldMappings(fields);
 * // { titleField: 'Title', startDateField: 'EventDate', ... }
 * ```
 */
export function autoDetectFieldMappings(fields: IFieldInfo[]): {
  titleField: string;
  startDateField: string;
  endDateField: string;
  allDayField: string;
  categoryField: string;
  locationField: string;
} {
  // Build a lookup map for O(1) existence checks by internal name
  const fieldMap = new Map(fields.map(f => [f.internalName, f]));

  // --- Category detection ---
  // Prefer an exact "Category" column; otherwise scan for keyword matches
  let categoryField = '';
  if (fieldMap.has('Category')) categoryField = 'Category';
  else {
    for (const f of fields) {
      const name = (f.internalName + ' ' + f.displayName).toLowerCase();
      if ((f.fieldType === 'Choice' || f.fieldType === 'Text') &&
          (name.indexOf('category') >= 0 || name.indexOf('type') >= 0)) {
        categoryField = f.internalName;
        break;
      }
    }
  }

  // --- Location detection ---
  // Check well-known names first, then fall back to keyword scan
  let locationField = '';
  if (fieldMap.has('Location')) locationField = 'Location';
  else if (fieldMap.has('WorkAddress')) locationField = 'WorkAddress';
  else {
    for (const f of fields) {
      const name = (f.internalName + ' ' + f.displayName).toLowerCase();
      if (f.fieldType === 'Text' && (name.indexOf('location') >= 0 || name.indexOf('address') >= 0 || name.indexOf('venue') >= 0)) {
        locationField = f.internalName;
        break;
      }
    }
  }

  // --- Strategy 1: Standard OOB Events list schema ---
  if (fieldMap.has('EventDate') && fieldMap.has('EndDate')) {
    return {
      titleField: 'Title',
      startDateField: 'EventDate',
      endDateField: 'EndDate',
      allDayField: fieldMap.has('fAllDayEvent') ? 'fAllDayEvent' : '',
      categoryField,
      locationField: locationField || (fieldMap.has('Location') ? 'Location' : ''),
    };
  }

  // --- Strategy 2: Scan DateTime fields for recognizable names ---
  const dateFields = fields.filter(f =>
    f.fieldType === 'DateTime' || f.fieldType === 'Date'
  );

  let startField = '';
  let endField = '';

  for (const f of dateFields) {
    const name = f.internalName.toLowerCase();
    const title = f.displayName.toLowerCase();
    if (!startField && (name === 'start' || name === 'startdate' || name === 'start_x0020_date' ||
        title === 'start' || title === 'start date' || title === 'begin')) {
      startField = f.internalName;
    } else if (!endField && (name === 'end' || name === 'enddate' || name === 'end_x0020_date' ||
        title === 'end' || title === 'end date' || title === 'finish')) {
      endField = f.internalName;
    }
  }

  // --- Strategy 3: Fallback — use first two DateTime fields by ordinal position ---
  if (!startField && dateFields.length > 0) startField = dateFields[0].internalName;
  if (!endField && dateFields.length > 1) endField = dateFields[1].internalName;

  return {
    titleField: fieldMap.has('Title') ? 'Title' : '',
    startDateField: startField,
    endDateField: endField,
    allDayField: '',
    categoryField,
    locationField,
  };
}
