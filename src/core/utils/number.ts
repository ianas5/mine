/** Formats a kg value: integer as-is, otherwise one decimal (e.g. `80`, `82.5`). */
export function formatKg(value: number): string {
  return Number.isInteger(value) ? String(value) : value.toFixed(1);
}
