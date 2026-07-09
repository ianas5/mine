import { useEffect, useMemo, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { todayIso, type IsoDate } from '@/core/utils';
import { bodyRepository } from '@/data/repositories/bodyRepository';
import { nutritionRepository } from '@/data/repositories/nutritionRepository';
import { phaseRepository } from '@/data/repositories/phaseRepository';
import { settingsRepository } from '@/data/repositories/settingsRepository';
import { workoutRepository } from '@/data/repositories/workoutRepository';
import {
  computePhaseReport,
  type DailyNutrition,
  type PhaseReport,
  type TrainingWorkout,
  type WeighIn,
} from '@/domain/analytics';
import type { BodySnapshot } from '@/domain/body';
import type { Phase } from '@/domain/models';

/** All declared phases, most recent first. Reactive to the `phases` table. */
export function usePhases(): readonly Phase[] | undefined {
  const version = useTableVersion('phases');
  const [phases, setPhases] = useState<readonly Phase[] | undefined>(undefined);
  useEffect(() => {
    let live = true;
    void phaseRepository.listPhases().then((rows) => {
      if (live) setPhases(rows);
    });
    return () => {
      live = false;
    };
  }, [version]);
  return phases;
}

interface PhaseInputs {
  readonly workouts: readonly TrainingWorkout[];
  readonly weighIns: readonly WeighIn[];
  readonly snapshots: readonly BodySnapshot[];
  readonly nutritionDays: readonly DailyNutrition[];
  readonly weeklyWorkoutTarget: number;
  readonly defaultBodyweightKg: number | null;
  readonly heightCm: number | null;
}

/** The full-history inputs a Phase Report windows over. Fetched once and shared; a phase
 * is a lens over these, never a filter applied at query time. Reactive to every source. */
function usePhaseInputs(): PhaseInputs | undefined {
  const version = useTableVersion('workouts', 'body', 'nutrition', 'settings');
  const [inputs, setInputs] = useState<PhaseInputs | undefined>(undefined);
  useEffect(() => {
    let live = true;
    void (async () => {
      const [workouts, weighIns, snapshots, nutritionDays, settings] = await Promise.all([
        workoutRepository.getTrainingWorkoutsSince('2000-01-01' as IsoDate),
        bodyRepository.getWeightLog(),
        bodyRepository.listSnapshots(),
        nutritionRepository.getDailyNutritionSince('2000-01-01' as IsoDate),
        settingsRepository.get(),
      ]);
      if (live) {
        setInputs({
          workouts,
          weighIns,
          snapshots,
          nutritionDays,
          weeklyWorkoutTarget: settings.weeklyWorkoutTarget,
          defaultBodyweightKg: settings.defaultBodyweightKg,
          heightCm: settings.heightCm,
        });
      }
    })();
    return () => {
      live = false;
    };
  }, [version]);
  return inputs;
}

/**
 * The Phase Report for one phase (`null` phaseId ⇒ the ongoing phase). `undefined` while
 * loading, `null` when no such phase exists. The report is computed purely from the shared
 * full-history inputs, so a completed block always reads the same (ANALYTICS §5.4).
 */
export function usePhaseReport(phaseId: string | null): PhaseReport | undefined | null {
  const phases = usePhases();
  const inputs = usePhaseInputs();
  return useMemo<PhaseReport | undefined | null>(() => {
    if (phases === undefined || inputs === undefined) return undefined;
    const phase =
      phaseId === null
        ? phases.find((p) => p.endDate === null)
        : phases.find((p) => p.id === phaseId);
    if (!phase) return null;
    return computePhaseReport({ phase, ...inputs, today: todayIso() });
  }, [phases, inputs, phaseId]);
}
