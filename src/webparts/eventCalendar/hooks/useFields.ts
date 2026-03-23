/**
 * @file useFields.ts
 * @description React hook that fetches and caches the available SharePoint list
 *              fields (columns) for a given list. Used by the property pane to
 *              populate column-mapping dropdowns and the display-fields picker.
 * @author W. Kevin Wolff
 * @copyright TidalHealth
 */

import * as React from 'react';
import { IFieldInfo } from '../models/IFieldInfo';
import { fetchAllListFields } from '../services/FieldService';

/**
 * Shape returned by the {@link useFields} hook.
 */
export interface IUseFieldsResult {
  /** Array of user-facing fields retrieved from the selected list. */
  fields: IFieldInfo[];
  /** `true` while the field metadata request is in flight. */
  loading: boolean;
  /** Human-readable error message from the last failed fetch, or `undefined`. */
  error: string | undefined;
}

/**
 * Fetches the non-system fields for a SharePoint list and exposes them with
 * loading/error state.
 *
 * Re-fetches automatically whenever `listId` changes. When `listId` is
 * `undefined` (no list selected yet), the hook resets to an empty field array
 * without making a network request. A `cancelled` flag in the effect cleanup
 * guards against stale responses after the list selection changes mid-flight.
 *
 * @param listId - GUID of the SharePoint list, or `undefined` if none is selected.
 * @returns An {@link IUseFieldsResult} with the field array, loading flag, and error message.
 *
 * @example
 * ```tsx
 * const { fields, loading, error } = useFields(props.selectedListId);
 * if (loading) return <Spinner />;
 * ```
 */
export function useFields(listId: string | undefined): IUseFieldsResult {
  const [fields, setFields] = React.useState<IFieldInfo[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>();

  React.useEffect(() => {
    // No list selected — clear fields and skip the request
    if (!listId) {
      setFields([]);
      return;
    }

    // Cancellation flag — prevents stale async results from updating state
    let cancelled = false;
    setLoading(true);
    setError(undefined);

    fetchAllListFields(listId)
      .then((result: IFieldInfo[]) => {
        if (!cancelled) {
          setFields(result);
          setLoading(false);
        }
      })
      .catch((err: { message?: string }) => {
        if (!cancelled) {
          setError(err.message || 'Failed to load fields');
          setLoading(false);
        }
      });

    // Cleanup: mark this effect cycle as cancelled so in-flight fetches are ignored
    return () => { cancelled = true; };
  }, [listId]);

  return { fields, loading, error };
}
