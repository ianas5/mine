import { useState, type ReactNode } from 'react';
import { KeyboardAvoidingView, Modal, Platform, Pressable, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { useTheme } from '@/core/theme';

import { Dialog } from './Dialog';

interface SheetProps {
  readonly visible: boolean;
  readonly onClose: () => void;
  readonly children: ReactNode;
  readonly title?: string | undefined;
  /** When true, dismissal asks "Discard entry?" once (UI_UX §2 dirty-state guard). */
  readonly dirty?: boolean;
}

/** Bottom sheet for all logging flows — grabber, lg top radius, keyboard-safe. */
export function Sheet(props: SheetProps): ReactNode {
  const theme = useTheme();
  const insets = useSafeAreaInsets();
  const [confirmingDiscard, setConfirmingDiscard] = useState(false);

  const requestClose = (): void => {
    if (props.dirty === true) {
      setConfirmingDiscard(true);
      return;
    }
    props.onClose();
  };

  return (
    <Modal visible={props.visible} transparent animationType="slide" onRequestClose={requestClose}>
      <View style={{ flex: 1, backgroundColor: 'rgba(0, 0, 0, 0.5)' }}>
        <Pressable style={{ flex: 1 }} onPress={requestClose} accessibilityLabel="Dismiss sheet" />
        <KeyboardAvoidingView behavior={Platform.OS === 'ios' ? 'padding' : undefined}>
          <View
            style={{
              backgroundColor: theme.color.surface,
              borderTopLeftRadius: theme.radius.lg,
              borderTopRightRadius: theme.radius.lg,
              paddingHorizontal: theme.space.lg,
              paddingTop: theme.space.sm,
              paddingBottom: insets.bottom + theme.space.lg,
            }}
            accessibilityViewIsModal
          >
            <View style={{ alignItems: 'center', paddingVertical: theme.space.sm }}>
              <View
                style={{
                  width: 36,
                  height: 4,
                  borderRadius: theme.radius.full,
                  backgroundColor: theme.color.border,
                }}
                accessibilityLabel="Sheet grabber"
              />
            </View>
            {props.title !== undefined ? (
              <Text
                style={{
                  ...theme.type.heading,
                  color: theme.color.textPrimary,
                  marginBottom: theme.space.md,
                }}
              >
                {props.title}
              </Text>
            ) : null}
            {props.children}
          </View>
        </KeyboardAvoidingView>
      </View>
      <Dialog
        visible={confirmingDiscard}
        title="Discard entry?"
        message="Your unsaved input will be lost."
        confirmLabel="Discard"
        onConfirm={() => {
          setConfirmingDiscard(false);
          props.onClose();
        }}
        onCancel={() => setConfirmingDiscard(false)}
      />
    </Modal>
  );
}
