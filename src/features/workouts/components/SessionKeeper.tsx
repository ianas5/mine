import { useEffect } from 'react';
import { AppState } from 'react-native';

import { workoutRepository } from '@/data/repositories/workoutRepository';

import { recoverSession } from '../logic/sessionRecovery';
import { serializeSession } from '../schemas/sessionDraftSchema';
import { useSessionStore, type SessionState } from '../stores/useSessionStore';

/** Value edits (weight/reps/name) coalesce; structural changes persist immediately. */
const DEBOUNCE_MS = 600;

/**
 * A cheap fingerprint of the session's high-value structure — exercise count and
 * completed-set count. When it changes (exercise added, set completed) we persist
 * immediately; lesser edits (typing a weight) merely debounce. This keeps the
 * most important events crash-safe at the instant they happen.
 */
function structuralSignature(state: SessionState): string {
  let cells = 0;
  let done = 0;
  for (const ex of state.exercises) {
    cells += ex.sets.length;
    for (const s of ex.sets) if (s.done) done += 1;
  }
  return `${state.active ? 1 : 0}:${state.exercises.length}:${cells}:${done}`;
}

/**
 * The crash-safety engine (ARCHITECTURE §7.1), mounted once at the composition
 * root inside the DB gate. Runs recovery on launch, then checkpoints the live
 * session to the SQLite draft — immediately on structural changes, debounced on
 * value edits, and flushed on app-background. On finish/discard the draft is
 * removed so a completed or abandoned session never re-appears. Renders nothing.
 */
export function SessionKeeper(): null {
  useEffect(() => {
    void recoverSession();
  }, []);

  useEffect(() => {
    let timer: ReturnType<typeof setTimeout> | null = null;
    let lastSignature = structuralSignature(useSessionStore.getState());

    const clearTimer = (): void => {
      if (timer !== null) {
        clearTimeout(timer);
        timer = null;
      }
    };

    const flush = (): void => {
      clearTimer();
      const state = useSessionStore.getState();
      if (state.active) {
        void workoutRepository.checkpointDraft(serializeSession(state));
      }
    };

    const unsubscribe = useSessionStore.subscribe((state, prev) => {
      // Session ended (finish or discard): make sure no draft lingers. Harmless if
      // finish already deleted it transactionally.
      if (prev.active && !state.active) {
        clearTimer();
        lastSignature = structuralSignature(state);
        void workoutRepository.discardDraft();
        return;
      }
      if (!state.active) {
        lastSignature = structuralSignature(state);
        return;
      }

      const signature = structuralSignature(state);
      if (signature !== lastSignature) {
        lastSignature = signature;
        flush(); // exercise added / set completed → persist now
      } else {
        clearTimer();
        timer = setTimeout(flush, DEBOUNCE_MS); // value edit → coalesce
      }
    });

    const appState = AppState.addEventListener('change', (status) => {
      if (status === 'background' || status === 'inactive') flush();
    });

    return () => {
      unsubscribe();
      appState.remove();
      clearTimer();
    };
  }, []);

  return null;
}
