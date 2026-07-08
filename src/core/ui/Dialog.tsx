import type { ReactNode } from 'react';
import { Modal, Pressable, Text, View } from 'react-native';

import { triggerHaptic, useTheme } from '@/core/theme';

import { Button } from './Button';

interface DialogProps {
  readonly visible: boolean;
  readonly title: string;
  readonly message?: string;
  readonly confirmLabel: string;
  readonly cancelLabel?: string;
  readonly onConfirm: () => void;
  readonly onCancel: () => void;
}

/** Destructive-confirm dialog — the ONLY sanctioned use of Dialog (DESIGN_SYSTEM §6, P10). */
export function Dialog(props: DialogProps): ReactNode {
  const theme = useTheme();
  return (
    <Modal visible={props.visible} transparent animationType="fade" onRequestClose={props.onCancel}>
      <Pressable
        onPress={props.onCancel}
        style={{
          flex: 1,
          backgroundColor: 'rgba(0, 0, 0, 0.5)',
          alignItems: 'center',
          justifyContent: 'center',
          padding: theme.space.xl,
        }}
        accessibilityLabel="Dismiss dialog"
      >
        <Pressable
          onPress={(e) => e.stopPropagation()}
          style={{
            width: '100%',
            maxWidth: 340,
            backgroundColor: theme.color.surfaceRaised,
            borderRadius: theme.radius.lg,
            borderWidth: 1,
            borderColor: theme.color.border,
            padding: theme.space.xl,
            gap: theme.space.md,
          }}
          accessibilityViewIsModal
        >
          <Text style={{ ...theme.type.heading, color: theme.color.textPrimary }}>
            {props.title}
          </Text>
          {props.message !== undefined ? (
            <Text style={{ ...theme.type.body, color: theme.color.textSecondary }}>
              {props.message}
            </Text>
          ) : null}
          <View style={{ gap: theme.space.sm, marginTop: theme.space.sm }}>
            <Button
              variant="destructive"
              size="md"
              label={props.confirmLabel}
              onPress={() => {
                triggerHaptic('warning');
                props.onConfirm();
              }}
            />
            <Button
              variant="ghost"
              size="md"
              label={props.cancelLabel ?? 'Cancel'}
              onPress={props.onCancel}
            />
          </View>
        </Pressable>
      </Pressable>
    </Modal>
  );
}
