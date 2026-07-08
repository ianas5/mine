import { useState, type ReactNode } from 'react';
import { View } from 'react-native';
import { Settings } from 'lucide-react-native';

import { useTheme } from '@/core/theme';
import { Button, Chip, IconButton, Input, SegmentedControl, Section, Stepper } from '@/core/ui';

/** Action & input primitives: Button, IconButton, Input, Stepper, Chip, SegmentedControl. */
export function GalleryActions(): ReactNode {
  const theme = useTheme();
  const [weightKg, setWeightKg] = useState(82.5);
  const [reps, setReps] = useState(8);
  const [slot, setSlot] = useState('Breakfast');
  const [range, setRange] = useState(1);
  const [text, setText] = useState('');
  const [numeric, setNumeric] = useState('82.4');

  return (
    <>
      <Section title="Button">
        <View style={{ gap: theme.space.sm }}>
          <Button label="Primary lg" onPress={() => undefined} />
          <Button label="Secondary md" variant="secondary" size="md" onPress={() => undefined} />
          <Button label="Ghost" variant="ghost" size="md" onPress={() => undefined} />
          <Button label="Destructive" variant="destructive" size="md" onPress={() => undefined} />
          <Button label="Disabled" disabled onPress={() => undefined} />
          <Button label="Loading" loading onPress={() => undefined} />
        </View>
      </Section>
      <Section title="IconButton">
        <View style={{ flexDirection: 'row', gap: theme.space.sm }}>
          <IconButton
            icon={<Settings color={theme.color.textPrimary} size={24} strokeWidth={1.75} />}
            onPress={() => undefined}
            accessibilityLabel="Settings"
          />
          <IconButton
            icon={<Settings color={theme.color.textPrimary} size={24} strokeWidth={1.75} />}
            onPress={() => undefined}
            accessibilityLabel="Settings disabled"
            disabled
          />
        </View>
      </Section>
      <Section title="Input">
        <View style={{ gap: theme.space.md }}>
          <Input label="Notes" value={text} onChangeText={setText} placeholder="Optional note…" />
          <Input
            label="Weight"
            value={numeric}
            onChangeText={setNumeric}
            keyboardType="decimal-pad"
            unit="KG"
          />
          <Input
            label="With error"
            value=""
            onChangeText={() => undefined}
            error="Weight must be between 0 and 500 kg"
          />
        </View>
      </Section>
      <Section title="Stepper">
        <View style={{ gap: theme.space.md, alignItems: 'flex-start' }}>
          <Stepper
            value={weightKg}
            onChange={setWeightKg}
            step={2.5}
            min={0}
            format={(v) => `${v} kg`}
            accessibilityLabel="Weight stepper"
          />
          <Stepper
            value={reps}
            onChange={setReps}
            step={1}
            min={0}
            max={100}
            format={(v) => `${v} reps`}
            accessibilityLabel="Reps stepper"
          />
        </View>
      </Section>
      <Section title="Chip">
        <View style={{ flexDirection: 'row', gap: theme.space.sm, flexWrap: 'wrap' }}>
          {['Breakfast', 'Lunch', 'Dinner', 'Snacks'].map((label) => (
            <Chip
              key={label}
              label={label}
              selected={slot === label}
              onPress={() => setSlot(label)}
            />
          ))}
        </View>
      </Section>
      <Section title="SegmentedControl">
        <SegmentedControl
          options={['7D', '30D', '3M', '6M', '1Y', 'All']}
          selectedIndex={range}
          onChange={setRange}
          accessibilityLabel="Time range"
        />
      </Section>
    </>
  );
}
