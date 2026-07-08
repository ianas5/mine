import type { ReactNode } from 'react';
import { ScrollView, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { useTheme } from '@/core/theme';

interface ScreenProps {
  readonly children: ReactNode;
  /** Wrap content in a ScrollView. Default false. */
  readonly scroll?: boolean;
  /** Apply the standard horizontal gutter (space.lg). Default true. */
  readonly gutter?: boolean;
}

/** Safe-area screen wrapper — every screen uses it (DESIGN_SYSTEM §6). */
export function Screen(props: ScreenProps): ReactNode {
  const theme = useTheme();
  const insets = useSafeAreaInsets();
  const base = {
    flex: 1,
    backgroundColor: theme.color.bg,
    paddingTop: insets.top,
  } as const;
  const gutterPad = props.gutter === false ? 0 : theme.space.lg;

  if (props.scroll) {
    return (
      <View style={base}>
        <ScrollView
          contentContainerStyle={{
            paddingHorizontal: gutterPad,
            paddingBottom: insets.bottom + theme.space.xxl,
          }}
          keyboardShouldPersistTaps="handled"
        >
          {props.children}
        </ScrollView>
      </View>
    );
  }
  return <View style={{ ...base, paddingHorizontal: gutterPad }}>{props.children}</View>;
}
