import { workoutRepository } from '@/data/repositories/workoutRepository';

import { parseSessionDraft } from '../schemas/sessionDraftSchema';
import { useSessionStore } from '../stores/useSessionStore';

export type RecoveryOutcome = 'restored' | 'discarded' | 'none';

/**
 * Crash recovery, run once after the DB is ready (ARCHITECTURE §7.1). Loads the
 * draft, validates it, and silently rehydrates the session so the workout is
 * simply *still there* — recovery is invisible (the session bar reappears and the
 * banner offers an explicit Resume/Discard). A corrupt, empty, or legacy draft is
 * discarded gracefully rather than surfacing an error. Never clobbers a session
 * that is already live in memory.
 */
export async function recoverSession(): Promise<RecoveryOutcome> {
  const payload = await workoutRepository.loadDraft();
  if (payload === null) return 'none';

  const draft = parseSessionDraft(payload);
  if (draft === null || draft.exercises.length === 0) {
    await workoutRepository.discardDraft();
    return 'discarded';
  }

  if (useSessionStore.getState().active) return 'none';

  useSessionStore.getState().actions.restore(draft);
  return 'restored';
}
