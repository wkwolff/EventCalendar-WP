import * as React from 'react';
import { IFieldInfo } from '../models/IFieldInfo';
import { fetchAllListFields } from '../services/FieldService';

export interface IUseFieldsResult {
  fields: IFieldInfo[];
  loading: boolean;
  error: string | undefined;
}

export function useFields(listId: string | undefined): IUseFieldsResult {
  const [fields, setFields] = React.useState<IFieldInfo[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>();

  React.useEffect(() => {
    if (!listId) {
      setFields([]);
      return;
    }

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

    return () => { cancelled = true; };
  }, [listId]);

  return { fields, loading, error };
}
