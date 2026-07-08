import { migrate } from 'drizzle-orm/expo-sqlite/migrator';
import { useEffect, useState, type ReactNode } from 'react';
import { Text, View } from 'react-native';

import { initDb } from './client';

type GateState = 'pending' | 'ready' | 'error';

/** The drizzle-kit runtime bundle shape (data/schema/migrations/migrations.js). */
export type MigrationBundle = Parameters<typeof migrate>[1];

interface DbGateProps {
  /** Injected by the composition root (app/_layout) from data/schema — core never imports data. */
  readonly migrations: MigrationBundle;
  /** Idempotent seeder run after migrations, before ready (DATABASE §5.6). Injected too. */
  readonly afterMigrate?: () => Promise<void>;
  readonly children: ReactNode;
}

/**
 * The DB-ready gate (ARCHITECTURE §8/§12): opens the database and applies
 * pending migrations before any feature renders. Splash stays visible while
 * pending; a migration failure renders a calm error state, never a white screen.
 */
export function DbGate(props: DbGateProps): ReactNode {
  const [state, setState] = useState<GateState>('pending');
  const [message, setMessage] = useState('');

  useEffect(() => {
    let live = true;
    (async () => {
      try {
        const db = initDb();
        // The gate only ever constructs the expo driver; the cast narrows the union.
        await migrate(db as Parameters<typeof migrate>[0], props.migrations);
        await props.afterMigrate?.();
        if (live) setState('ready');
      } catch (cause) {
        if (live) {
          setMessage(cause instanceof Error ? cause.message : String(cause));
          setState('error');
        }
      }
    })();
    return () => {
      live = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps -- migrations bundle is static
  }, []);

  if (state === 'pending') {
    return null; // splash screen remains visible
  }
  if (state === 'error') {
    return (
      <View style={{ flex: 1, alignItems: 'center', justifyContent: 'center', gap: 8 }}>
        <Text accessibilityRole="alert">Database failed to open.</Text>
        <Text>{message}</Text>
      </View>
    );
  }
  return props.children;
}
