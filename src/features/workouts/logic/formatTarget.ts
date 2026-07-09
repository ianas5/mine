import type { SessionTarget } from '../stores/useSessionStore';

/** Formats a rep range: `8–10`, `8` (equal), `8+` (min only), `≤10` (max only). */
function formatReps(repMin: number | null, repMax: number | null): string | null {
  if (repMin !== null && repMax !== null) {
    return repMin === repMax ? `${repMin}` : `${repMin}–${repMax}`;
  }
  if (repMin !== null) return `${repMin}+`;
  if (repMax !== null) return `≤${repMax}`;
  return null;
}

/**
 * A planned target as a compact logging-time hint (UI_UX §5.2), e.g.
 * `Target · 3 × 8–10 @ RPE 8`. Returns null when the template set no targets — a
 * plan is guidance, never fabricated numbers (P8).
 */
export function formatTarget(target: SessionTarget | null): string | null {
  if (target === null) return null;

  const reps = formatReps(target.repMin, target.repMax);
  const parts: string[] = [];
  if (target.sets !== null && reps !== null) parts.push(`${target.sets} × ${reps}`);
  else if (target.sets !== null) parts.push(`${target.sets} sets`);
  else if (reps !== null) parts.push(`${reps} reps`);
  if (target.rpe !== null) parts.push(`@ RPE ${target.rpe}`);

  return parts.length > 0 ? `Target · ${parts.join(' ')}` : null;
}
