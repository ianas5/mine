/**
 * Plausibility ranges and formula constants — the single source for these
 * numbers (CODING_STANDARDS §6.2). Values from FITNESS_DOMAIN §8.13 and §3.5.
 * A threshold literal appearing inline anywhere else is a defect.
 */
export const LOAD_KG_MIN = 0;
export const LOAD_KG_MAX = 1000;
export const REPS_MIN = 0;
export const REPS_MAX = 100;
export const RPE_MIN = 0;
export const RPE_MAX = 10;
export const RIR_MIN = 0;
export const RIR_MAX = 10;

/** Epley e1RM is trusted for PRs only at or below this rep count (FITNESS_DOMAIN §3.5). */
export const E1RM_MAX_TRUSTED_REPS = 12;
