import { create } from 'zustand';

/**
 * Default rest between working sets when the exercise has no remembered value
 * (UI_UX §3.2/§4.2: "exercise's rest_seconds, else 90 s default"). Per-exercise
 * `rest_seconds` arrives with templates in a later phase; until then the store
 * remembers any in-session extension per exercise.
 */
export const DEFAULT_REST_SEC = 90;

/** Extend granularity for the +time control. */
export const REST_EXTEND_SEC = 30;

interface RestActions {
  /**
   * Auto-started when a working set is completed (UI_UX §4.2). Wall-clock based.
   * `defaultSec` seeds a template's `rest_seconds` when this exercise has no
   * remembered in-session extension; falls back to the 90 s default.
   */
  readonly start: (exerciseLocalId: string, now: number, defaultSec?: number | null) => void;
  /** Add time to the running rest and remember it as this exercise's preference. */
  readonly extend: (deltaSec: number, now: number) => void;
  /** User skipped the rest (or it elapsed) — clears the countdown, keeps prefs. */
  readonly skip: () => void;
  /** Full teardown on session finish/discard — also forgets per-exercise prefs. */
  readonly reset: () => void;
}

export interface RestTimerState {
  readonly running: boolean;
  /** Wall-clock ms at which rest ends; survives background/foreground honestly. */
  readonly endsAt: number | null;
  readonly exerciseLocalId: string | null;
  readonly durationSec: number;
  /** Remembered rest length per exercise localId, so extensions carry to next set. */
  readonly prefs: Readonly<Record<string, number>>;
  readonly actions: RestActions;
}

const INITIAL = {
  running: false,
  endsAt: null as number | null,
  exerciseLocalId: null as string | null,
  durationSec: DEFAULT_REST_SEC,
  prefs: {} as Record<string, number>,
};

export const useRestTimerStore = create<RestTimerState>((set, get) => ({
  ...INITIAL,
  actions: {
    start: (exerciseLocalId, now, defaultSec) => {
      const durationSec =
        get().prefs[exerciseLocalId] ??
        (defaultSec != null && defaultSec > 0 ? defaultSec : DEFAULT_REST_SEC);
      set({ running: true, endsAt: now + durationSec * 1000, exerciseLocalId, durationSec });
    },

    extend: (deltaSec, now) =>
      set((state) => {
        if (!state.running || state.endsAt === null) return state;
        const durationSec = Math.max(0, state.durationSec + deltaSec);
        // Extend from whichever is later — now or the current end — so adding time
        // after the timer already lapsed still gives a full delta.
        const endsAt = Math.max(now, state.endsAt) + deltaSec * 1000;
        const prefs =
          state.exerciseLocalId !== null
            ? { ...state.prefs, [state.exerciseLocalId]: durationSec }
            : state.prefs;
        return { ...state, running: true, endsAt, durationSec, prefs };
      }),

    skip: () => set({ running: false, endsAt: null, exerciseLocalId: null }),

    reset: () => set({ ...INITIAL, prefs: {} }),
  },
}));

export const useRestActions = (): RestActions => useRestTimerStore((state) => state.actions);
