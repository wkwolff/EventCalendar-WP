import { getSP } from './PnPSetup';
import { IFieldInfo } from '../models/IFieldInfo';

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

// Field types that cannot be queried via REST $select
const UNSUPPORTED_FIELD_TYPES = new Set([
  'Geolocation',
]);

// Fields with known REST API issues
const UNSUPPORTED_FIELD_NAMES = new Set([
  'ParticipantsPicker',
  'ParticipantsPickerId',
]);

/** Fetch ALL non-system fields for a list (used for column mapping dropdowns) */
export async function fetchAllListFields(listId: string): Promise<IFieldInfo[]> {
  const sp = getSP();
  const fields = await sp.web.lists.getById(listId).fields
    .filter("Hidden eq false")
    .select('InternalName', 'Title', 'TypeAsString', 'Required')();

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

/** Fetch display fields (excludes the mapped core fields) */
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

/** Auto-detect standard event list fields, returns suggested mappings */
export function autoDetectFieldMappings(fields: IFieldInfo[]): {
  titleField: string;
  startDateField: string;
  endDateField: string;
  allDayField: string;
  categoryField: string;
  locationField: string;
} {
  const fieldMap = new Map(fields.map(f => [f.internalName, f]));

  // Detect category field
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

  // Detect location field
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

  // Standard Events list fields
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

  // Try common custom field names
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

  // Fallback: use first two date fields
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
