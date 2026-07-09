import { z } from 'zod';

import { LOAD_TYPES } from '@/domain/fitness';

import type { SessionExercise, SessionState } from '../stores/useSessionStore';

/**
 * Crash-safe draft of the active session (ARCHITECTURE §7.1, DATABASE §3.4).
 * The payload is written on every meaningful mutation and on background, then
 * Zod-validated before it is ever trusted on resume — an unparseable or
 * out-of-shape draft is discarded gracefully rather than crashing the app.
 *
 * `localId`s are intentionally NOT persisted: they are ephemeral in-memory keys
 * regenerated on restore, so a recovered session can never collide with ids
 * minted afterward.
 */
const UNILATERAL_COUNTING = ['none', 'single_doubled', 'per_side'] as const;

const draftSetSchema = z.object({
  weightKg: z.number().finite().nonnegative(),
  reps: z.number().int().nonnegative(),
  rpe: z.number().finite().nonnegative().nullable(),
  warmup: z.boolean(),
  done: z.boolean(),
});

const draftTargetSchema = z.object({
  sets: z.number().int().nonnegative().nullable(),
  repMin: z.number().int().nonnegative().nullable(),
  repMax: z.number().int().nonnegative().nullable(),
  rpe: z.number().finite().nonnegative().nullable(),
});

const draftExerciseSchema = z.object({
  exerciseId: z.string().min(1),
  name: z.string(),
  loadType: z.enum(LOAD_TYPES),
  unilateralCounting: z.enum(UNILATERAL_COUNTING),
  // Optional so a draft written by an earlier build still parses (recovery must
  // never fail on a shape gap); absent → treated as no target / no rest.
  target: draftTargetSchema.nullish(),
  restSeconds: z.number().int().nonnegative().nullish(),
  sets: z.array(draftSetSchema),
});

export const sessionDraftSchema = z.object({
  version: z.literal(1),
  name: z.string(),
  startedAt: z.number().finite(),
  templateId: z.string().nullish(),
  exercises: z.array(draftExerciseSchema),
});

/** The validated recovery snapshot (what `restore` rehydrates from). */
export type SessionDraft = z.infer<typeof sessionDraftSchema>;

const CURRENT_VERSION = 1 as const;

/** Serializes a live session into the draft payload string (checkpoint write). */
export function serializeSession(state: SessionState): string {
  const draft: SessionDraft = {
    version: CURRENT_VERSION,
    name: state.name,
    startedAt: state.startedAt ?? 0,
    templateId: state.templateId,
    exercises: state.exercises.map((ex: SessionExercise) => ({
      exerciseId: ex.exerciseId,
      name: ex.name,
      loadType: ex.loadType,
      unilateralCounting: ex.unilateralCounting,
      target: ex.target,
      restSeconds: ex.restSeconds,
      sets: ex.sets.map((s) => ({
        weightKg: s.weightKg,
        reps: s.reps,
        rpe: s.rpe,
        warmup: s.warmup,
        done: s.done,
      })),
    })),
  };
  return JSON.stringify(draft);
}

/**
 * Parses and validates a draft payload. Returns the snapshot, or `null` for any
 * malformed/corrupt/legacy-version payload so recovery can discard it silently.
 */
export function parseSessionDraft(payload: string): SessionDraft | null {
  let raw: unknown;
  try {
    raw = JSON.parse(payload);
  } catch {
    return null;
  }
  const result = sessionDraftSchema.safeParse(raw);
  return result.success ? result.data : null;
}
