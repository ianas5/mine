/**
 * The body-metric field set (FITNESS_DOMAIN §5.1), keyed to the `body_snapshots`
 * columns. Order is the natural top-to-bottom measuring sequence (UI_UX §4.5):
 * composition first, then circumferences head-to-toe, each bilateral site per side.
 */
export const BODY_FIELDS = [
  'weightKg',
  'bodyFatPct',
  'muscleMassKg',
  'visceralFat',
  'bmi',
  'neckCm',
  'chestCm',
  'waistCm',
  'hipsCm',
  'leftArmCm',
  'rightArmCm',
  'leftForearmCm',
  'rightForearmCm',
  'leftThighCm',
  'rightThighCm',
  'leftCalfCm',
  'rightCalfCm',
] as const;

export type BodyField = (typeof BODY_FIELDS)[number];

export type BodyFieldGroup = 'composition' | 'circumference';

export interface BodyFieldMeta {
  readonly key: BodyField;
  readonly label: string;
  readonly unit: string;
  readonly group: BodyFieldGroup;
  /** Stepper step for the entry control (0.1 kg for weight, 0.1 cm for tape, etc.). */
  readonly step: number;
}

export const BODY_FIELD_META: Record<BodyField, BodyFieldMeta> = {
  weightKg: { key: 'weightKg', label: 'Weight', unit: 'kg', group: 'composition', step: 0.1 },
  bodyFatPct: { key: 'bodyFatPct', label: 'Body fat', unit: '%', group: 'composition', step: 0.1 },
  muscleMassKg: {
    key: 'muscleMassKg',
    label: 'Muscle mass',
    unit: 'kg',
    group: 'composition',
    step: 0.1,
  },
  visceralFat: {
    key: 'visceralFat',
    label: 'Visceral fat',
    unit: '',
    group: 'composition',
    step: 1,
  },
  bmi: { key: 'bmi', label: 'BMI', unit: '', group: 'composition', step: 0.1 },
  neckCm: { key: 'neckCm', label: 'Neck', unit: 'cm', group: 'circumference', step: 0.1 },
  chestCm: { key: 'chestCm', label: 'Chest', unit: 'cm', group: 'circumference', step: 0.1 },
  waistCm: { key: 'waistCm', label: 'Waist', unit: 'cm', group: 'circumference', step: 0.1 },
  hipsCm: { key: 'hipsCm', label: 'Hips', unit: 'cm', group: 'circumference', step: 0.1 },
  leftArmCm: { key: 'leftArmCm', label: 'Left arm', unit: 'cm', group: 'circumference', step: 0.1 },
  rightArmCm: {
    key: 'rightArmCm',
    label: 'Right arm',
    unit: 'cm',
    group: 'circumference',
    step: 0.1,
  },
  leftForearmCm: {
    key: 'leftForearmCm',
    label: 'Left forearm',
    unit: 'cm',
    group: 'circumference',
    step: 0.1,
  },
  rightForearmCm: {
    key: 'rightForearmCm',
    label: 'Right forearm',
    unit: 'cm',
    group: 'circumference',
    step: 0.1,
  },
  leftThighCm: {
    key: 'leftThighCm',
    label: 'Left thigh',
    unit: 'cm',
    group: 'circumference',
    step: 0.1,
  },
  rightThighCm: {
    key: 'rightThighCm',
    label: 'Right thigh',
    unit: 'cm',
    group: 'circumference',
    step: 0.1,
  },
  leftCalfCm: {
    key: 'leftCalfCm',
    label: 'Left calf',
    unit: 'cm',
    group: 'circumference',
    step: 0.1,
  },
  rightCalfCm: {
    key: 'rightCalfCm',
    label: 'Right calf',
    unit: 'cm',
    group: 'circumference',
    step: 0.1,
  },
};
