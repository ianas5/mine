import { useState, type ReactNode } from 'react';
import { Animated, Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';

interface SegmentedControlProps {
  readonly options: readonly string[];
  readonly selectedIndex: number;
  readonly onChange: (index: number) => void;
  readonly accessibilityLabel?: string;
}

/** The single either/or + time-range switcher — sliding thumb at motion.fast. */
export function SegmentedControl(props: SegmentedControlProps): ReactNode {
  const theme = useTheme();
  const [trackWidth, setTrackWidth] = useState(0);
  const [thumbX] = useState(() => new Animated.Value(0));
  const segmentWidth = props.options.length > 0 ? trackWidth / props.options.length : 0;

  const moveThumb = (index: number): void => {
    Animated.timing(thumbX, {
      toValue: index * segmentWidth,
      duration: theme.motion.fast,
      useNativeDriver: true,
    }).start();
  };

  return (
    <View
      accessibilityRole="tablist"
      accessibilityLabel={props.accessibilityLabel}
      onLayout={(e) => {
        const width = e.nativeEvent.layout.width;
        setTrackWidth(width);
        thumbX.setValue(props.selectedIndex * (width / props.options.length));
      }}
      style={{
        flexDirection: 'row',
        backgroundColor: theme.color.surface,
        borderRadius: theme.radius.sm,
        borderWidth: 1,
        borderColor: theme.color.border,
        overflow: 'hidden',
      }}
    >
      {segmentWidth > 0 ? (
        <Animated.View
          style={{
            position: 'absolute',
            top: 0,
            bottom: 0,
            width: segmentWidth,
            backgroundColor: theme.color.accentSoft,
            transform: [{ translateX: thumbX }],
          }}
        />
      ) : null}
      {props.options.map((option, index) => {
        const selected = index === props.selectedIndex;
        return (
          <Pressable
            key={option}
            onPress={() => {
              moveThumb(index);
              props.onChange(index);
            }}
            accessibilityRole="tab"
            accessibilityLabel={option}
            accessibilityState={{ selected }}
            style={{
              flex: 1,
              alignItems: 'center',
              paddingVertical: theme.space.sm,
              minHeight: 36,
              justifyContent: 'center',
            }}
          >
            <Text
              style={{
                ...theme.type.caption,
                color: selected ? theme.color.accent : theme.color.textSecondary,
              }}
            >
              {option}
            </Text>
          </Pressable>
        );
      })}
    </View>
  );
}
