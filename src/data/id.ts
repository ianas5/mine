import * as Crypto from 'expo-crypto';

/** App-generated stable UUID with a type prefix (DATABASE §2.1). */
export function newId(prefix: string): string {
  return `${prefix}_${Crypto.randomUUID()}`;
}
