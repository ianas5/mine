import { create } from 'zustand';

import {
  LOAD_KG_MAX,
  REPS_MAX,
  RPE_MAX,
  type LoadType,
  type PreviewSet,
  type UnilateralCounting,
} from '@/domain/fitness';
import type { Exercise } from '@/domain/models';

import type { SessionDraft } from '../schemas/sessionDraftSchema';

export interface SessionSet {
  readonly localId: string;
  readonly weightKg: number;
  readonly reps: number;
  readonly rpe: number | null;
  readonly warmup: boolean;
  readonly done: boolean;
}

export interface SessionExercise {
  readonly localId: string;
  readonly exerciseId: string;
  readonly name: string;
  readonly loadType: LoadType;
  readonly unilateralCounting: UnilateralCounting;
  readonly sets: readonly SessionSet[];
}

export interface SessionState {
  readonly active: boolean;
  /** True only for a session rehydrated from a crash draft — drives the recovery banner. */
  readonly recovered: boolean;
  readonly name: string;
  readonly startedAt: number | null;
  readonly exercises: readonly SessionExercise[];
  readonly actions: SessionActions;
}

interface SetPatch {
  readonly weightKg?: number;
  readonly reps?: number;
  readonly rpe?: number | null;
  readonly warmup?: boolean;
}

interface SessionActions {
  readonly start: (startedAt: number, name?: string) => void;
  readonly restore: (draft: SessionDraft) => void;
  readonly acknowledgeRecovery: () => void;
  readonly setName: (name: string) => void;
  readonly addExercise: (exercise: Exercise, prefill?: readonly PreviewSet[]) => void;
  readonly removeExercise: (exerciseLocalId: string) => void;
  readonly setCounting: (exerciseLocalId: string, counting: UnilateralCounting) => void;
  readonly addSet: (exerciseLocalId: string) => void;
  readonly updateSet: (exerciseLocalId: string, setLocalId: string, patch: SetPatch) => void;
  readonly toggleSetDone: (exerciseLocalId: string, setLocalId: string) => void;
  readonly removeSet: (exerciseLocalId: string, setLocalId: string) => void;
  readonly discard: () => void;
}

let seq = 0;
const localId = (prefix: string): string => `${prefix}${(seq += 1)}`;

const clamp = (value: number, min: number, max: number): number =>
  Math.min(Math.max(value, min), max);

const INITIAL = {
  active: false,
  recovered: false,
  name: '',
  startedAt: null as number | null,
  exercises: [],
};

function mapExercise(
  state: SessionState,
  exerciseLocalId: string,
  fn: (exercise: SessionExercise) => SessionExercise,
): SessionExercise[] {
  return state.exercises.map((ex) => (ex.localId === exerciseLocalId ? fn(ex) : ex));
}

export const useSessionStore = create<SessionState>((set, get) => ({
  ...INITIAL,
  actions: {
    start: (startedAt, name = 'Workout') =>
      set({ active: true, recovered: false, name, startedAt, exercises: [] }),

    // Rehydrate a crash-recovered session (ARCHITECTURE §7.1). localIds are minted
    // fresh through the same counter so ids created after recovery never collide.
    restore: (draft) =>
      set({
        active: true,
        recovered: true,
        name: draft.name,
        startedAt: draft.startedAt,
        exercises: draft.exercises.map((ex) => ({
          localId: localId('ex'),
          exerciseId: ex.exerciseId,
          name: ex.name,
          loadType: ex.loadType,
          unilateralCounting: ex.unilateralCounting,
          sets: ex.sets.map((s) => ({
            localId: localId('set'),
            weightKg: s.weightKg,
            reps: s.reps,
            rpe: s.rpe,
            warmup: s.warmup,
            done: s.done,
          })),
        })),
      }),

    // Dismiss the recovery banner without ending the session (the workout stays live).
    acknowledgeRecovery: () => set({ recovered: false }),

    setName: (name) => set({ name }),

    // Prefill from last time's working sets (Phase 5) so a repeated exercise is
    // logged with zero typing — tap ✓ down the list. Falls back to one empty set.
    addExercise: (exercise, prefill) =>
      set((state) => {
        const source = prefill && prefill.length > 0 ? prefill : [{ weightKg: 0, reps: 0 }];
        return {
          exercises: [
            ...state.exercises,
            {
              localId: localId('ex'),
              exerciseId: exercise.id,
              name: exercise.name,
              loadType: exercise.loadType,
              unilateralCounting: exercise.defaultUnilateral ? 'single_doubled' : 'none',
              sets: source.map((s) => ({
                localId: localId('set'),
                weightKg: s.weightKg,
                reps: s.reps,
                rpe: null,
                warmup: false,
                done: false,
              })),
            },
          ],
        };
      }),

    removeExercise: (exerciseLocalId) =>
      set((state) => ({
        exercises: state.exercises.filter((ex) => ex.localId !== exerciseLocalId),
      })),

    setCounting: (exerciseLocalId, counting) =>
      set((state) => ({
        exercises: mapExercise(state, exerciseLocalId, (ex) => ({
          ...ex,
          unilateralCounting: counting,
        })),
      })),

    // New sets inherit the previous set's weight/reps within THIS session (not
    // cross-workout history — that is Phase 5). Enables ≤1-tap repeat sets.
    addSet: (exerciseLocalId) =>
      set((state) => ({
        exercises: mapExercise(state, exerciseLocalId, (ex) => {
          const prev = ex.sets[ex.sets.length - 1];
          return {
            ...ex,
            sets: [
              ...ex.sets,
              {
                localId: localId('set'),
                weightKg: prev?.weightKg ?? 0,
                reps: prev?.reps ?? 0,
                rpe: null,
                warmup: false,
                done: false,
              },
            ],
          };
        }),
      })),

    updateSet: (exerciseLocalId, setLocalId, patch) =>
      set((state) => ({
        exercises: mapExercise(state, exerciseLocalId, (ex) => ({
          ...ex,
          sets: ex.sets.map((s) =>
            s.localId === setLocalId
              ? {
                  ...s,
                  ...(patch.weightKg !== undefined && {
                    weightKg: clamp(patch.weightKg, 0, LOAD_KG_MAX),
                  }),
                  ...(patch.reps !== undefined && { reps: clamp(patch.reps, 0, REPS_MAX) }),
                  ...(patch.rpe !== undefined && {
                    rpe: patch.rpe === null ? null : clamp(patch.rpe, 0, RPE_MAX),
                  }),
                  ...(patch.warmup !== undefined && { warmup: patch.warmup }),
                }
              : s,
          ),
        })),
      })),

    toggleSetDone: (exerciseLocalId, setLocalId) =>
      set((state) => ({
        exercises: mapExercise(state, exerciseLocalId, (ex) => ({
          ...ex,
          sets: ex.sets.map((s) => (s.localId === setLocalId ? { ...s, done: !s.done } : s)),
        })),
      })),

    removeSet: (exerciseLocalId, setLocalId) =>
      set((state) => ({
        exercises: mapExercise(state, exerciseLocalId, (ex) => ({
          ...ex,
          sets: ex.sets.filter((s) => s.localId !== setLocalId),
        })),
      })),

    discard: () => set({ ...INITIAL }),
  },
}));

export const useSessionActions = (): SessionActions => useSessionStore((state) => state.actions);
