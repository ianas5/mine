import { useRouter } from 'expo-router';
import { ArrowLeft, ImageOff } from 'lucide-react-native';
import { useMemo, useState, type ReactNode } from 'react';
import { Image, Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import {
  Card,
  Chip,
  EmptyState,
  IconButton,
  Screen,
  SegmentedControl,
  Sheet,
  Skeleton,
} from '@/core/ui';
import { formatRelativeDate } from '@/core/utils';
import { PHOTO_ANGLES, PHOTO_ANGLE_LABELS, type PhotoAngle } from '@/domain/photos';

import { usePhotos } from '../hooks/usePhotos';
import type { PhotoWithStatus } from '@/data/repositories/photoRepository';

const MODES = ['Side by side', 'Before / After'] as const;

/** Compare two progress photos (UI_UX §5.2) — same angle, side-by-side or an A/B toggle. */
export function PhotoCompareScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const photos = usePhotos();

  const anglesWithPhotos = useMemo(
    () => PHOTO_ANGLES.filter((a) => (photos ?? []).some((p) => p.angle === a)),
    [photos],
  );
  const [angle, setAngle] = useState<PhotoAngle | null>(null);
  const activeAngle = angle ?? anglesWithPhotos[0] ?? null;

  // Photos for the active angle, oldest → newest (so A defaults to the earliest).
  const forAngle = useMemo(
    () =>
      (photos ?? [])
        .filter((p) => p.angle === activeAngle)
        .slice()
        .sort((x, y) => (x.date < y.date ? -1 : x.date > y.date ? 1 : 0)),
    [photos, activeAngle],
  );

  const [pickA, setPickA] = useState<string | null>(null);
  const [pickB, setPickB] = useState<string | null>(null);
  const [picking, setPicking] = useState<null | 'a' | 'b'>(null);
  const [mode, setMode] = useState(0);
  const [showAfter, setShowAfter] = useState(true);

  const photoA = forAngle.find((p) => p.id === pickA) ?? forAngle[0] ?? null;
  const photoB = forAngle.find((p) => p.id === pickB) ?? forAngle[forAngle.length - 1] ?? null;

  const enoughForAngle = forAngle.length >= 2;

  return (
    <Screen scroll>
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          gap: theme.space.sm,
          marginTop: theme.space.sm,
          marginBottom: theme.space.lg,
        }}
      >
        <IconButton
          icon={<ArrowLeft color={theme.color.textPrimary} size={24} strokeWidth={1.75} />}
          onPress={() => router.back()}
          accessibilityLabel="Back"
        />
        <Text style={{ ...theme.type.title, color: theme.color.textPrimary, flex: 1 }}>
          Compare Photos
        </Text>
      </View>

      {photos === undefined ? (
        <View style={{ gap: theme.space.md }}>
          <Skeleton height={48} />
          <Skeleton height={320} />
        </View>
      ) : anglesWithPhotos.length === 0 ? (
        <EmptyState title="No photos yet to compare." />
      ) : (
        <View style={{ gap: theme.space.lg }}>
          <View style={{ flexDirection: 'row', gap: theme.space.xs }}>
            {anglesWithPhotos.map((a) => (
              <Chip
                key={a}
                label={PHOTO_ANGLE_LABELS[a]}
                selected={activeAngle === a}
                onPress={() => {
                  setAngle(a);
                  setPickA(null);
                  setPickB(null);
                }}
              />
            ))}
          </View>

          {!enoughForAngle ? (
            <EmptyState title="Capture at least two of this angle to compare them." />
          ) : (
            <>
              <SegmentedControl
                options={MODES as unknown as string[]}
                selectedIndex={mode}
                onChange={setMode}
                accessibilityLabel="Compare mode"
              />

              {mode === 0 ? (
                <View style={{ flexDirection: 'row', gap: theme.space.sm }}>
                  <ComparePane label="Before" photo={photoA} onPress={() => setPicking('a')} />
                  <ComparePane label="After" photo={photoB} onPress={() => setPicking('b')} />
                </View>
              ) : (
                <BeforeAfter
                  photo={showAfter ? photoB : photoA}
                  showingAfter={showAfter}
                  onToggle={() => setShowAfter((v) => !v)}
                />
              )}

              {mode === 0 ? (
                <Text
                  style={{
                    ...theme.type.caption,
                    color: theme.color.textTertiary,
                    textAlign: 'center',
                  }}
                >
                  Tap a photo to change the date.
                </Text>
              ) : (
                <View style={{ flexDirection: 'row', gap: theme.space.sm }}>
                  <PickField label="Before" photo={photoA} onPress={() => setPicking('a')} />
                  <PickField label="After" photo={photoB} onPress={() => setPicking('b')} />
                </View>
              )}
            </>
          )}
        </View>
      )}

      <Sheet
        visible={picking !== null}
        onClose={() => setPicking(null)}
        title={picking === 'a' ? 'Before photo' : 'After photo'}
      >
        <View>
          {forAngle.map((photo) => (
            <Pressable
              key={photo.id}
              onPress={() => {
                if (picking === 'a') setPickA(photo.id);
                else setPickB(photo.id);
                setPicking(null);
              }}
              accessibilityRole="button"
              accessibilityLabel={`${photo.date}`}
              style={({ pressed }) => ({
                paddingVertical: theme.space.md,
                opacity: pressed ? 0.6 : 1,
              })}
            >
              <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
                {formatRelativeDate(photo.date)} · {photo.date}
              </Text>
            </Pressable>
          ))}
        </View>
      </Sheet>
    </Screen>
  );
}

function ComparePane(props: {
  readonly label: string;
  readonly photo: PhotoWithStatus | null;
  readonly onPress: () => void;
}): ReactNode {
  const theme = useTheme();
  const { photo } = props;
  return (
    <Pressable
      onPress={props.onPress}
      accessibilityRole="button"
      accessibilityLabel={`${props.label}${photo ? `, ${photo.date}` : ''}`}
      style={({ pressed }) => ({ flex: 1, gap: theme.space.xs, opacity: pressed ? 0.8 : 1 })}
    >
      <Text style={{ ...theme.type.micro, color: theme.color.textTertiary }}>
        {props.label.toUpperCase()}
      </Text>
      {photo && !photo.fileMissing ? (
        <Image
          source={{ uri: photo.uri }}
          resizeMode="cover"
          style={{ width: '100%', aspectRatio: 3 / 4, borderRadius: theme.radius.md }}
        />
      ) : (
        <MissingPane />
      )}
      <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
        {photo ? formatRelativeDate(photo.date) : '—'}
      </Text>
    </Pressable>
  );
}

function BeforeAfter(props: {
  readonly photo: PhotoWithStatus | null;
  readonly showingAfter: boolean;
  readonly onToggle: () => void;
}): ReactNode {
  const theme = useTheme();
  const { photo } = props;
  return (
    <Pressable
      onPress={props.onToggle}
      accessibilityRole="button"
      accessibilityLabel={`Showing ${props.showingAfter ? 'after' : 'before'}, tap to toggle`}
      style={({ pressed }) => ({ opacity: pressed ? 0.9 : 1 })}
    >
      {photo && !photo.fileMissing ? (
        <Image
          source={{ uri: photo.uri }}
          resizeMode="cover"
          style={{ width: '100%', aspectRatio: 3 / 4, borderRadius: theme.radius.md }}
        />
      ) : (
        <View style={{ aspectRatio: 3 / 4 }}>
          <MissingPane />
        </View>
      )}
      <View
        style={{
          position: 'absolute',
          top: theme.space.sm,
          left: theme.space.sm,
          paddingHorizontal: theme.space.sm,
          paddingVertical: theme.space.xs,
          borderRadius: theme.radius.sm,
          backgroundColor: theme.color.surfaceRaised,
        }}
      >
        <Text style={{ ...theme.type.caption, color: theme.color.textPrimary }}>
          {props.showingAfter ? 'After' : 'Before'}
          {photo ? ` · ${formatRelativeDate(photo.date)}` : ''}
        </Text>
      </View>
    </Pressable>
  );
}

function PickField(props: {
  readonly label: string;
  readonly photo: PhotoWithStatus | null;
  readonly onPress: () => void;
}): ReactNode {
  const theme = useTheme();
  return (
    <Pressable
      onPress={props.onPress}
      accessibilityRole="button"
      accessibilityLabel={`Choose ${props.label} date`}
      style={({ pressed }) => ({ flex: 1, opacity: pressed ? 0.6 : 1 })}
    >
      <Card style={{ gap: theme.space.xs }}>
        <Text style={{ ...theme.type.micro, color: theme.color.textTertiary }}>
          {props.label.toUpperCase()}
        </Text>
        <Text style={{ ...theme.type.caption, color: theme.color.textPrimary }}>
          {props.photo ? formatRelativeDate(props.photo.date) : '—'}
        </Text>
      </Card>
    </Pressable>
  );
}

function MissingPane(): ReactNode {
  const theme = useTheme();
  return (
    <View
      style={{
        width: '100%',
        height: '100%',
        aspectRatio: 3 / 4,
        borderRadius: theme.radius.md,
        borderWidth: 1,
        borderColor: theme.color.border,
        backgroundColor: theme.color.surface,
        alignItems: 'center',
        justifyContent: 'center',
      }}
    >
      <ImageOff color={theme.color.textTertiary} size={20} />
    </View>
  );
}
