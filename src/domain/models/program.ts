import type { EpochMs } from '@/core/utils';
import type { LoadType } from '@/domain/fitness';

/** Planned targets for one exercise within a template (DATABASE §3.3). Never performance. */
export interface TemplateTarget {
  readonly sets: number | null;
  readonly repMin: number | null;
  readonly repMax: number | null;
  readonly rpe: number | null;
  readonly restSeconds: number | null;
}

/** One exercise line of a template, with its resolved catalog details. */
export interface TemplateExercise {
  readonly id: string;
  readonly exerciseId: string;
  readonly name: string;
  readonly loadType: LoadType;
  readonly defaultUnilateral: boolean;
  readonly position: number;
  readonly target: TemplateTarget;
  readonly notes: string | null;
}

/** A single-session blueprint (DATABASE §3.3). A planning tool — not history. */
export interface Template {
  readonly id: string;
  readonly programId: string | null;
  readonly name: string;
  readonly position: number;
  /** Scheduled days (0 = Monday … 6 = Sunday); empty when unscheduled. A session may
   * repeat on several weekdays. Always sorted ascending, de-duplicated. */
  readonly weekdays: readonly number[];
  readonly notes: string | null;
  readonly exercises: readonly TemplateExercise[];
}

/** A named collection of session templates (DATABASE §3.3). At most one is active. */
export interface Program {
  readonly id: string;
  readonly name: string;
  readonly notes: string | null;
  readonly isActive: boolean;
  readonly isArchived: boolean;
  readonly createdAt: EpochMs;
  readonly updatedAt: EpochMs;
  readonly templates: readonly Template[];
}
