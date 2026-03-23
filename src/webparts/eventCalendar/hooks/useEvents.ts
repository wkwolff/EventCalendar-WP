import * as React from 'react';
import { IEventItem } from '../models/IEventItem';
import { fetchEvents, IEventFieldMapping } from '../services/EventService';

export interface IUseEventsResult {
  events: IEventItem[];
  loading: boolean;
  error: string | undefined;
  refresh: () => void;
}

export function useEvents(
  listId: string | undefined,
  fieldMapping: IEventFieldMapping,
  selectedFields: string[],
  maxEvents: number
): IUseEventsResult {
  const [events, setEvents] = React.useState<IEventItem[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>();
  const [refreshKey, setRefreshKey] = React.useState(0);

  const fieldsKey = selectedFields.join(',');
  const mappingKey = fieldMapping.titleField + '|' + fieldMapping.startDateField + '|' +
    fieldMapping.endDateField + '|' + fieldMapping.allDayField + '|' +
    fieldMapping.categoryField + '|' + fieldMapping.locationField;

  React.useEffect(() => {
    if (!listId || !fieldMapping.startDateField) {
      setEvents([]);
      return;
    }

    let cancelled = false;
    setLoading(true);
    setError(undefined);

    fetchEvents(listId, fieldMapping, selectedFields, maxEvents)
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

    return () => { cancelled = true; };
  }, [listId, mappingKey, fieldsKey, maxEvents, refreshKey]);

  const refresh = React.useCallback(() => setRefreshKey(k => k + 1), []);

  return { events, loading, error, refresh };
}
