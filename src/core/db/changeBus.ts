import { useSyncExternalStore } from 'react';

/**
 * Minimal change-bus (ARCHITECTURE §7): every repository write announces which
 * table(s) changed; data hooks re-query on their tables only. Deliberately just
 * an emitter keyed by table name — no query keys, no staleness policies.
 */
export type TableName =
  'settings' | 'exercises' | 'workouts' | 'programs' | 'nutrition' | 'body' | 'photos' | 'phases';

type Listener = () => void;

const versions = new Map<TableName, number>();
const listeners = new Map<TableName, Set<Listener>>();

/** Called by repositories as the last step of every write (CODING_STANDARDS §8.3). */
export function emitTableChanges(...tables: readonly TableName[]): void {
  for (const table of tables) {
    versions.set(table, (versions.get(table) ?? 0) + 1);
    for (const listener of listeners.get(table) ?? []) {
      listener();
    }
  }
}

/** Subscribe to writes on specific tables. Returns an unsubscribe function. */
export function subscribeToTables(tables: readonly TableName[], listener: Listener): () => void {
  for (const table of tables) {
    let set = listeners.get(table);
    if (!set) {
      set = new Set();
      listeners.set(table, set);
    }
    set.add(listener);
  }
  return () => {
    for (const table of tables) {
      listeners.get(table)?.delete(listener);
    }
  };
}

/**
 * The shared subscription helper for data hooks (CODING_STANDARDS §4.4):
 * returns a number that changes whenever any of the given tables is written —
 * use it as an effect dependency to re-query.
 */
export function useTableVersion(...tables: readonly TableName[]): number {
  return useSyncExternalStore(
    (onChange) => subscribeToTables(tables, onChange),
    () => tables.reduce((sum, table) => sum + (versions.get(table) ?? 0), 0),
  );
}
