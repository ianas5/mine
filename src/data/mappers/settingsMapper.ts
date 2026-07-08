import type { Settings } from '@/domain/models';

import type { SettingsRow } from '../schema/tables';

export function rowToSettings(row: SettingsRow): Settings {
  return {
    weeklyWorkoutTarget: row.weeklyWorkoutTarget,
    defaultBodyweightKg: row.defaultBodyweightKg,
    heightCm: row.heightCm,
    waterCupMl: row.waterCupMl,
  };
}
