import { useRouter } from 'expo-router';
import { ArrowLeft, ArrowLeftRight, ImageOff } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Image, Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Dialog, EmptyState, IconButton, Screen, Sheet, Skeleton } from '@/core/ui';
import { formatRelativeDate } from '@/core/utils';
import { PHOTO_ANGLE_LABELS, groupPhotosByDate, oldestMissingAngle } from '@/domain/photos';

import { AddPhotoSheet } from '../components/AddPhotoSheet';
import { usePhotos } from '../hooks/usePhotos';
import { photoRepository, type PhotoWithStatus } from '@/data/repositories/photoRepository';

/** Progress photos timeline (UI_UX §4.5/§5.2) — grouped by date, newest first. */
export function PhotosScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const photos = usePhotos();
  const [adding, setAdding] = useState(false);
  const [viewing, setViewing] = useState<PhotoWithStatus | null>(null);
  const [confirmDelete, setConfirmDelete] = useState(false);

  const groups = photos ? groupPhotosByDate(photos) : [];
  const defaultAngle = oldestMissingAngle(photos ?? []);
  const canCompare = (photos?.length ?? 0) >= 2;

  const remove = async (): Promise<void> => {
    if (viewing) await photoRepository.deletePhoto(viewing.id);
    setConfirmDelete(false);
    setViewing(null);
  };

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
          Progress Photos
        </Text>
        {canCompare ? (
          <IconButton
            icon={<ArrowLeftRight color={theme.color.textSecondary} size={22} strokeWidth={1.75} />}
            onPress={() => router.push('/measurements/photos/compare')}
            accessibilityLabel="Compare photos"
          />
        ) : null}
      </View>

      <View style={{ marginBottom: theme.space.lg }}>
        <Button label="Add Photo" onPress={() => setAdding(true)} />
      </View>

      {photos === undefined ? (
        <View style={{ gap: theme.space.md }}>
          <Skeleton height={160} />
          <Skeleton height={160} />
        </View>
      ) : groups.length === 0 ? (
        <EmptyState title="No photos yet — tap Add Photo to capture your first." />
      ) : (
        <View style={{ gap: theme.space.lg }}>
          {groups.map((group) => (
            <View key={group.date} style={{ gap: theme.space.sm }}>
              <Text style={{ ...theme.type.bodyStrong, color: theme.color.textSecondary }}>
                {formatRelativeDate(group.date)}
              </Text>
              <View style={{ flexDirection: 'row', flexWrap: 'wrap', gap: theme.space.sm }}>
                {group.photos.map((photo) => (
                  <Thumbnail key={photo.id} photo={photo} onPress={() => setViewing(photo)} />
                ))}
              </View>
            </View>
          ))}
        </View>
      )}

      <AddPhotoSheet
        visible={adding}
        defaultAngle={defaultAngle}
        onClose={() => setAdding(false)}
      />

      <Sheet
        visible={viewing !== null}
        onClose={() => setViewing(null)}
        title={viewing ? `${PHOTO_ANGLE_LABELS[viewing.angle]} · ${viewing.date}` : ''}
      >
        {viewing ? (
          <View style={{ gap: theme.space.lg }}>
            {viewing.fileMissing ? (
              <MissingTile height={360} />
            ) : (
              <Image
                source={{ uri: viewing.uri }}
                accessibilityLabel={`${PHOTO_ANGLE_LABELS[viewing.angle]} photo, ${viewing.date}`}
                resizeMode="cover"
                style={{ width: '100%', height: 360, borderRadius: theme.radius.md }}
              />
            )}
            <Button
              label="Delete photo"
              variant="destructive"
              onPress={() => setConfirmDelete(true)}
            />
          </View>
        ) : null}
      </Sheet>

      <Dialog
        visible={confirmDelete}
        title="Delete this photo?"
        message="The photo and its file are removed permanently."
        confirmLabel="Delete"
        onConfirm={() => void remove()}
        onCancel={() => setConfirmDelete(false)}
      />
    </Screen>
  );
}

function Thumbnail(props: {
  readonly photo: PhotoWithStatus;
  readonly onPress: () => void;
}): ReactNode {
  const theme = useTheme();
  const { photo } = props;
  return (
    <Pressable
      onPress={props.onPress}
      accessibilityRole="button"
      accessibilityLabel={`${PHOTO_ANGLE_LABELS[photo.angle]} photo, ${photo.date}`}
      style={({ pressed }) => ({ width: '31.5%', opacity: pressed ? 0.7 : 1 })}
    >
      {photo.fileMissing ? (
        <View style={{ aspectRatio: 3 / 4 }}>
          <MissingTile />
        </View>
      ) : (
        <Image
          source={{ uri: photo.uri }}
          resizeMode="cover"
          style={{ width: '100%', aspectRatio: 3 / 4, borderRadius: theme.radius.md }}
        />
      )}
      <View
        style={{
          position: 'absolute',
          left: theme.space.xs,
          bottom: theme.space.xs,
          paddingHorizontal: theme.space.xs,
          paddingVertical: theme.space.xs,
          borderRadius: theme.radius.sm,
          backgroundColor: theme.color.surfaceRaised,
        }}
      >
        <Text style={{ ...theme.type.micro, color: theme.color.textPrimary }}>
          {PHOTO_ANGLE_LABELS[photo.angle]}
        </Text>
      </View>
    </Pressable>
  );
}

/** Placeholder for a row whose file went missing (never silently hidden). */
function MissingTile(props: { readonly height?: number }): ReactNode {
  const theme = useTheme();
  return (
    <View
      style={{
        width: '100%',
        height: props.height ?? '100%',
        aspectRatio: props.height ? undefined : 3 / 4,
        borderRadius: theme.radius.md,
        borderWidth: 1,
        borderColor: theme.color.border,
        backgroundColor: theme.color.surface,
        alignItems: 'center',
        justifyContent: 'center',
        gap: theme.space.xs,
      }}
    >
      <ImageOff color={theme.color.textTertiary} size={22} />
      <Text style={{ ...theme.type.micro, color: theme.color.textTertiary }}>File missing</Text>
    </View>
  );
}
