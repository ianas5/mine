/** Domain configuration — SQLite-backed, in backups (ARCHITECTURE §6, DATABASE §3.1). */
export interface Settings {
  /** FITNESS_DOMAIN §3.8 — planned sessions per week when no program schedule exists. */
  readonly weeklyWorkoutTarget: number;
  /** Fallback bodyweight for bodyweight-load exercises (FITNESS_DOMAIN §3.4). Null = unset. */
  readonly defaultBodyweightKg: number | null;
  /** For derived BMI (FITNESS_DOMAIN §5.2). Null = unset. */
  readonly heightCm: number | null;
  /** Water logging increment (FITNESS_DOMAIN §4.1). */
  readonly waterCupMl: number;
}
