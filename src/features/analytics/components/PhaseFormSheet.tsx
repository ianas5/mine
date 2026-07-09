import { useState, type ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Chip, Input, Sheet } from '@/core/ui';
import { todayIso } from '@/core/utils';
import {
  PhaseValidationError,
  phaseRepository,
  type PhaseValidationReason,
} from '@/data/repositories/phaseRepository';
import { PHASE_TYPES, PHASE_TYPE_LABELS, type Phase, type PhaseType } from '@/domain/models';

const ISO_RE = /^\d{4}-\d{2}-\d{2}$/;
const isValidIso = (s: string): boolean => ISO_RE.test(s) && !Number.isNaN(Date.parse(s));

function validationMessage(reason: PhaseValidationReason): string {
  return reason === 'overlap'
    ? 'This overlaps an existing phase. End your current phase first, or adjust the dates.'
    : 'The end date must be on or after the start date.';
}

interface Props {
  readonly visible: boolean;
  readonly onClose: () => void;
  /** When set, the sheet edits this phase; otherwise it declares a new one. */
  readonly editing?: Phase | null;
  readonly onSaved?: () => void;
}

/** Declare or edit a training phase (ANALYTICS §5.4). Dates are entered as YYYY-MM-DD so a
 * block can be declared over historical data; the repository enforces the no-overlap rule. */
export function PhaseFormSheet(props: Props): ReactNode {
  const theme = useTheme();
  const editing = props.editing ?? null;

  const [name, setName] = useState(() => editing?.name ?? '');
  const [type, setType] = useState<PhaseType>(() => editing?.type ?? 'cutting');
  const [start, setStart] = useState(() => editing?.startDate ?? todayIso());
  const [ongoing, setOngoing] = useState(() => (editing ? editing.endDate === null : true));
  const [end, setEnd] = useState(() => editing?.endDate ?? '');
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);

  const startError = start.length > 0 && !isValidIso(start) ? 'Use YYYY-MM-DD' : undefined;
  const endError = !ongoing && end.length > 0 && !isValidIso(end) ? 'Use YYYY-MM-DD' : undefined;

  const save = async (): Promise<void> => {
    if (!isValidIso(start)) {
      setError('Enter a valid start date.');
      return;
    }
    if (!ongoing && !isValidIso(end)) {
      setError('Enter a valid end date, or mark the phase ongoing.');
      return;
    }
    const endDate = ongoing ? null : end;
    setSaving(true);
    setError(null);
    try {
      if (editing) {
        await phaseRepository.updatePhase(editing.id, { name, type, startDate: start, endDate });
      } else {
        await phaseRepository.createPhase({ name, type, startDate: start, endDate, notes: null });
      }
      props.onSaved?.();
      props.onClose();
    } catch (e) {
      if (e instanceof PhaseValidationError) setError(validationMessage(e.reason));
      else setError('Could not save this phase.');
    } finally {
      setSaving(false);
    }
  };

  return (
    <Sheet
      visible={props.visible}
      onClose={props.onClose}
      title={editing ? 'Edit phase' : 'Declare a phase'}
    >
      <View style={{ gap: theme.space.lg }}>
        <Input label="Name" value={name} onChangeText={setName} placeholder="e.g. Winter Cut" />

        <View style={{ gap: theme.space.sm }}>
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>Type</Text>
          <View style={{ flexDirection: 'row', flexWrap: 'wrap', gap: theme.space.sm }}>
            {PHASE_TYPES.map((t) => (
              <Chip
                key={t}
                label={PHASE_TYPE_LABELS[t]}
                selected={type === t}
                onPress={() => setType(t)}
              />
            ))}
          </View>
        </View>

        <Input
          label="Start date"
          value={start}
          onChangeText={setStart}
          placeholder="YYYY-MM-DD"
          keyboardType="numbers-and-punctuation"
          error={startError}
        />

        <View style={{ flexDirection: 'row', gap: theme.space.sm }}>
          <Chip label="Still ongoing" selected={ongoing} onPress={() => setOngoing(true)} />
          <Chip label="Has ended" selected={!ongoing} onPress={() => setOngoing(false)} />
        </View>

        {!ongoing ? (
          <Input
            label="End date"
            value={end}
            onChangeText={setEnd}
            placeholder="YYYY-MM-DD"
            keyboardType="numbers-and-punctuation"
            error={endError}
          />
        ) : null}

        {error !== null ? (
          <Text style={{ ...theme.type.caption, color: theme.color.danger }}>{error}</Text>
        ) : null}

        <Button
          label={editing ? 'Save phase' : 'Declare phase'}
          onPress={() => void save()}
          loading={saving}
        />
      </View>
    </Sheet>
  );
}
