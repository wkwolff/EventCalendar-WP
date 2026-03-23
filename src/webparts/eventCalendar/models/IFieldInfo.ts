/**
 * @file IFieldInfo.ts
 * @description Lightweight model representing a SharePoint list field (column).
 *              Used to populate the property-pane column-mapping dropdowns and
 *              the additional-fields picker.
 * @author W. Kevin Wolff
 * @copyright Wolff Creative
 */

/**
 * Metadata about a single SharePoint list field.
 *
 * This is a trimmed projection of the SP REST `SP.Field` resource, keeping
 * only the properties the web part needs for configuration UI and query building.
 *
 * @example
 * ```ts
 * const field: IFieldInfo = {
 *   internalName: 'EventDate',
 *   displayName: 'Start Time',
 *   fieldType: 'DateTime',
 *   required: true,
 * };
 * ```
 */
export interface IFieldInfo {
  /** The column's internal (static) name used in REST queries. */
  internalName: string;

  /** The human-readable column title shown in the SharePoint UI. */
  displayName: string;

  /** SharePoint field type string (e.g. "Text", "DateTime", "Choice"). */
  fieldType: string;

  /** Whether the column is marked as required on the list. */
  required: boolean;
}
