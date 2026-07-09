export {
  MUSCLE_GROUPS,
  MUSCLE_GROUP_LABELS,
  LOAD_TYPES,
  LOAD_TYPE_LABELS,
  MOVEMENT_PATTERN,
  BODY_SPLIT,
} from './taxonomy';
export type { MuscleGroup, LoadType, MovementPattern, BodySplit } from './taxonomy';
export {
  LOAD_KG_MIN,
  LOAD_KG_MAX,
  REPS_MIN,
  REPS_MAX,
  RPE_MIN,
  RPE_MAX,
  RIR_MIN,
  RIR_MAX,
  E1RM_MAX_TRUSTED_REPS,
} from './constants';
export {
  effectiveLoadKg,
  isWorkingSet,
  setVolumeKg,
  isVolumeLowConfidence,
  epleyOneRepMax,
  isE1rmEligible,
} from './sets';
export type { RawSet, UnilateralCounting } from './sets';
export { computeWorkoutStats, isCountableWorkout } from './workoutStats';
export type { StatExercise, WorkoutStats } from './workoutStats';
export { summarizeExerciseHistory } from './exerciseHistory';
export type { ExercisePreview, HistorySetRow, PreviewSet } from './exerciseHistory';
export { computeExerciseBests, detectNewPRs, EMPTY_BESTS } from './personalRecords';
export type { ExerciseBests, ExerciseSetRow, PrEvent, PrKind } from './personalRecords';
export { suggestTemplate } from './templateSuggestion';
export type { RecentTemplateUse, TemplateSuggestion } from './templateSuggestion';
