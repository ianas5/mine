import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import {
  latestFieldValues,
  weightLogWithDeltas,
  type BodyField,
  type BodySnapshot,
  type FieldLatest,
  type WeightLogEntry,
} from '@/domain/body';
import { bodyRepository } from '@/data/repositories/bodyRepository';

export interface BodyData {
  readonly snapshots: readonly BodySnapshot[];
  readonly latest: Record<BodyField, FieldLatest | null>;
  readonly weightLog: readonly WeightLogEntry[];
}

/** All body data for the Measurements home — snapshots, latest per field, weight log. */
export function useBodyData(): BodyData | undefined {
  const version = useTableVersion('body');
  const [data, setData] = useState<BodyData | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void Promise.all([bodyRepository.listSnapshots(), bodyRepository.getWeightLog()]).then(
      ([snapshots, weighIns]) => {
        if (!live) return;
        setData({
          snapshots,
          latest: latestFieldValues(snapshots),
          weightLog: weightLogWithDeltas(weighIns),
        });
      },
    );
    return () => {
      live = false;
    };
  }, [version]);

  return data;
}
