import { useState, type ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Chip, Sheet, showToast } from '@/core/ui';
import { todayIso } from '@/core/utils';
import { PHOTO_ANGLES, PHOTO_ANGLE_LABELS, type PhotoAngle } from '@/domain/photos';
import { photoRepository } from '@/data/repositories/photoRepository';

import { captureFromCamera, pickFromLibrary, type PickedImage } from '../logic/pickPhoto';

interface AddPhotoSheetProps {
  readonly visible: boolean;
  /** Smart default = the oldest missing angle (UI_UX §5.2). */
  readonly defaultAngle: PhotoAngle;
  readonly onClose: () => void;
}

/** Add Photo (UI_UX §4/§5.2) — angle pre-set to the oldest missing, pick or capture, save. */
export function AddPhotoSheet(props: AddPhotoSheetProps): ReactNode {
  const theme = useTheme();
  const [angle, setAngle] = useState<PhotoAngle>(props.defaultAngle);
  const [busy, setBusy] = useState(false);

  const save = async (pick: () => Promise<PickedImage | null>): Promise<void> => {
    setBusy(true);
    try {
      const picked = await pick();
      if (!picked) {
        setBusy(false);
        return;
      }
      await photoRepository.savePhoto({
        date: todayIso(),
        angle,
        sourceUri: picked.uri,
        width: picked.width,
        height: picked.height,
        notes: null,
      });
      showToast('Photo added', 'success');
      props.onClose();
    } catch {
      showToast('Could not add photo');
    } finally {
      setBusy(false);
    }
  };

  return (
    <Sheet visible={props.visible} onClose={props.onClose} title="Add Photo">
      <View style={{ gap: theme.space.lg }}>
        <View style={{ gap: theme.space.sm }}>
          <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>ANGLE</Text>
          <View style={{ flexDirection: 'row', gap: theme.space.xs }}>
            {PHOTO_ANGLES.map((a) => (
              <Chip
                key={a}
                label={PHOTO_ANGLE_LABELS[a]}
                selected={angle === a}
                onPress={() => setAngle(a)}
              />
            ))}
          </View>
        </View>

        <Button
          label="Choose from library"
          loading={busy}
          onPress={() => void save(pickFromLibrary)}
        />
        <Button
          label="Take photo"
          variant="secondary"
          disabled={busy}
          onPress={() => void save(captureFromCamera)}
        />
      </View>
    </Sheet>
  );
}
