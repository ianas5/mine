export {
  SERVING_UNITS,
  MEAL_SLOTS,
  MEAL_SLOT_LABELS,
  defaultSlotForHour,
  type ServingUnit,
  type MealSlot,
} from './taxonomy';
export {
  scalePortion,
  sumMacros,
  atwaterKcal,
  isKcalImplausible,
  ZERO_MACROS,
  type MacroSet,
  type FoodMacros,
} from './macros';
export {
  dayAdherence,
  remainingMacros,
  type TargetMacros,
  type MacroStatus,
  type DayAdherence,
} from './adherence';
