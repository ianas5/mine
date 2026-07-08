import { useRouter } from 'expo-router';
import { ArrowLeft } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme, useThemeControls } from '@/core/theme';
import {
  Card,
  IconButton,
  Input,
  Screen,
  Section,
  SegmentedControl,
  Skeleton,
  Stepper,
} from '@/core/ui';
import type { Settings } from '@/domain/models';

import { useSettings } from '../hooks/useSettings';

const THEME_OPTIONS = ['System', 'Dark', 'Light'] as const;
const THEME_VALUES = ['system', 'dark', 'light'] as const;

/** Rarely-visited configuration (UI_UX §3) — deliberately not a tab. */
export function SettingsScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const { override, setOverride } = useThemeControls();
  const { settings, update } = useSettings();

  return (
    <Screen scroll>
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          gap: theme.space.sm,
          marginTop: theme.space.sm,
          marginBottom: theme.space.xl,
        }}
      >
        <IconButton
          icon={<ArrowLeft color={theme.color.textPrimary} size={24} strokeWidth={1.75} />}
          onPress={() => router.back()}
          accessibilityLabel="Back"
        />
        <Text style={{ ...theme.type.title, color: theme.color.textPrimary }}>Settings</Text>
      </View>

      <Section title="Appearance">
        <SegmentedControl
          options={THEME_OPTIONS}
          selectedIndex={THEME_VALUES.indexOf(override)}
          onChange={(index) => setOverride(THEME_VALUES[index] ?? 'system')}
          accessibilityLabel="Theme"
        />
      </Section>

      {settings === null ? (
        <View style={{ gap: theme.space.md }}>
          <Skeleton height={52} />
          <Skeleton height={52} />
          <Skeleton height={52} />
        </View>
      ) : (
        <LoadedSettings settings={settings} update={update} />
      )}
    </Screen>
  );
}

interface LoadedSettingsProps {
  readonly settings: Settings;
  readonly update: (patch: Partial<Settings>) => void;
}

/** Rendered once settings exist so text fields can seed their state lazily (no sync effects). */
function LoadedSettings(props: LoadedSettingsProps): ReactNode {
  const theme = useTheme();
  const [heightText, setHeightText] = useState(() =>
    props.settings.heightCm === null ? '' : String(props.settings.heightCm),
  );
  const [bodyweightText, setBodyweightText] = useState(() =>
    props.settings.defaultBodyweightKg === null ? '' : String(props.settings.defaultBodyweightKg),
  );

  const commitNumber = (raw: string, field: 'heightCm' | 'defaultBodyweightKg'): void => {
    const parsed = Number.parseFloat(raw);
    props.update({ [field]: Number.isFinite(parsed) && parsed > 0 ? parsed : null });
  };

  return (
    <>
      <Section title="Training">
        <Card>
          <View
            style={{
              flexDirection: 'row',
              alignItems: 'center',
              justifyContent: 'space-between',
            }}
          >
            <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
              Weekly workout target
            </Text>
            <Stepper
              value={props.settings.weeklyWorkoutTarget}
              onChange={(next) => props.update({ weeklyWorkoutTarget: next })}
              step={1}
              min={1}
              max={14}
              accessibilityLabel="Weekly workout target"
            />
          </View>
        </Card>
      </Section>

      <Section title="Body">
        <View style={{ gap: theme.space.md }}>
          <Input
            label="Height"
            value={heightText}
            onChangeText={setHeightText}
            onBlur={() => commitNumber(heightText, 'heightCm')}
            keyboardType="decimal-pad"
            unit="CM"
            placeholder="—"
          />
          <Input
            label="Default bodyweight"
            value={bodyweightText}
            onChangeText={setBodyweightText}
            onBlur={() => commitNumber(bodyweightText, 'defaultBodyweightKg')}
            keyboardType="decimal-pad"
            unit="KG"
            placeholder="—"
          />
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
            Used for bodyweight exercises and derived BMI when set.
          </Text>
        </View>
      </Section>
    </>
  );
}
