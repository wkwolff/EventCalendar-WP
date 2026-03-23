/**
 * @file useEvents.ts
 * @description React hook that manages the lifecycle of fetching, caching, and
 *              refreshing calendar events from a SharePoint list. Wraps
 *              {@link fetchEvents} with state management, cancellation safety,
 *              and dependency-driven re-fetching.
 * @author W. Kevin Wolff
 * @copyright Wolff Creative
 */

import * as React from 'react';
import { IEventItem } from '../models/IEventItem';
import { IFieldInfo } from '../models/IFieldInfo';
import { fetchEvents, IEventFieldMapping } from '../services/EventService';

/**
 * Shape returned by the {@link useEvents} hook.
 */
export interface IUseEventsResult {
  /** The most recently fetched event array (empty while loading or on error). */
  events: IEventItem[];
  /** `true` while an async fetch is in flight. */
  loading: boolean;
  /** Human-readable error message from the last failed fetch, or `undefined`. */
  error: string | undefined;
  /** Imperatively trigger a re-fetch (e.g. after a manual data change). */
  refresh: () => void;
}

/**
 * Fetches and manages calendar events for the given list and mapping config.
 *
 * @param listId          - GUID of the SharePoint list, or `undefined` if none is selected yet.
 * @param fieldMapping    - Core column mapping configuration.
 * @param selectedFields  - Internal names of additional columns to fetch.
 * @param maxEvents       - Maximum number of events to retrieve.
 * @param availableFields - Field metadata used to detect User/Lookup fields for $expand.
 * @returns An {@link IUseEventsResult} with events, loading state, error, and a refresh callback.
 */
export function useEvents(
  listId: string | undefined,
  fieldMapping: IEventFieldMapping,
  selectedFields: string[],
  maxEvents: number,
  availableFields?: IFieldInfo[]
): IUseEventsResult {
  const [events, setEvents] = React.useState<IEventItem[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>();
  const [refreshKey, setRefreshKey] = React.useState(0);

  // Serialize arrays/objects into primitive strings so they can be used as
  // stable useEffect dependency values without causing infinite re-renders.
  const fieldsKey = selectedFields.join(',');
  const mappingKey = fieldMapping.titleField + '|' + fieldMapping.startDateField + '|' +
    fieldMapping.endDateField + '|' + fieldMapping.allDayField + '|' +
    fieldMapping.categoryField + '|' + fieldMapping.locationField;
  const availableFieldsKey = availableFields ? availableFields.map(f => f.internalName).join(',') : '';

  React.useEffect(() => {
    // Bail out early if the minimum required config is not yet available
    if (!listId || !fieldMapping.startDateField) {
      setEvents([]);
      return;
    }

    // Cancellation flag — prevents stale async results from updating state
    let cancelled = false;
    setLoading(true);
    setError(undefined);

    fetchEvents(listId, fieldMapping, selectedFields, maxEvents, availableFields)
      .then(result => {
        if (!cancelled) {
          setEvents(result);
          setLoading(false);
        }
      })
      .catch(err => {
        if (!cancelled) {
          setError(err.message || 'Failed to load events');
          setLoading(false);
        }
      });

    // Cleanup: mark this effect cycle as cancelled so in-flight fetches are ignored
    return () => { cancelled = true; };
  }, [listId, mappingKey, fieldsKey, maxEvents, refreshKey, availableFieldsKey]);

  /** Bumping the key forces useEffect to re-run and fetch fresh data. */
  const refresh = React.useCallback(() => setRefreshKey(k => k + 1), []);

  return { events, loading, error, refresh };
}
