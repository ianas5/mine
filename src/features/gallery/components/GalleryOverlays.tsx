import { useState, type ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Dialog, Input, Section, Sheet, showToast } from '@/core/ui';

/** Overlay & feedback primitives: Toast, Dialog, Sheet (incl. dirty-state guard). */
export function GalleryOverlays(): ReactNode {
  const theme = useTheme();
  const [dialogVisible, setDialogVisible] = useState(false);
  const [sheetVisible, setSheetVisible] = useState(false);
  const [sheetInput, setSheetInput] = useState('');

  return (
    <>
      <Section title="Toast">
        <View style={{ gap: theme.space.sm }}>
          <Button
            label="Show neutral toast"
            variant="secondary"
            size="md"
            onPress={() => showToast('Measurements saved')}
          />
          <Button
            label="Show success toast"
            variant="secondary"
            size="md"
            onPress={() => showToast('Workout saved · 3 PRs', 'success')}
          />
        </View>
      </Section>
      <Section title="Dialog (destructive only)">
        <Button
          label="Delete workout…"
          variant="destructive"
          size="md"
          onPress={() => setDialogVisible(true)}
        />
        <Dialog
          visible={dialogVisible}
          title="Delete workout?"
          message="This removes the session and its 14 sets. Records may recede."
          confirmLabel="Delete"
          onConfirm={() => setDialogVisible(false)}
          onCancel={() => setDialogVisible(false)}
        />
      </Section>
      <Section title="Sheet (with dirty guard)">
        <Button
          label="Open logging sheet"
          variant="secondary"
          size="md"
          onPress={() => setSheetVisible(true)}
        />
        <Sheet
          visible={sheetVisible}
          onClose={() => {
            setSheetVisible(false);
            setSheetInput('');
          }}
          title="Add Weight"
          dirty={sheetInput.length > 0}
        >
          <View style={{ gap: theme.space.md }}>
            <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
              Type something, then swipe down / tap outside to see the discard guard.
            </Text>
            <Input
              label="Weight"
              value={sheetInput}
              onChangeText={setSheetInput}
              keyboardType="decimal-pad"
              unit="KG"
            />
            <Button
              label="Save"
              onPress={() => {
                setSheetVisible(false);
                setSheetInput('');
                showToast('Weight saved');
              }}
            />
          </View>
        </Sheet>
      </Section>
    </>
  );
}
