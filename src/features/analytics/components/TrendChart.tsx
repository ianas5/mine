import { DashPathEffect } from '@shopify/react-native-skia';
import type { ReactNode } from 'react';
import { View } from 'react-native';
import { CartesianChart, Line } from 'victory-native';

import { useTheme } from '@/core/theme';
import type { SeriesPoint } from '@/domain/analytics';

interface TrendChartProps {
  /** Pre-bucketed series from the engine (ANALYTICS §4) — never resampled here. */
  readonly series: readonly SeriesPoint[];
  /** Optional dashed goal line (e.g. target weight). */
  readonly targetValue?: number | null;
  readonly height?: number;
}

interface Datum {
  readonly [key: string]: number;
  readonly x: number;
  readonly y: number;
  readonly target: number;
}

/**
 * The line plot inside a ChartFrame (DESIGN_SYSTEM §6.1), built on Victory Native XL
 * (Skia). Primary series = accent 2.5 pt; an optional dashed target reference. Consumes
 * engine-bucketed points as-is. Device-only rendering (Skia canvas) — never mounted in
 * tests; its host ChartFrame always carries the interpretation line.
 */
export function TrendChart(props: TrendChartProps): ReactNode {
  const theme = useTheme();
  const height = props.height ?? 200;
  const target = props.targetValue ?? null;

  // `target` is always present (falls back to y so the series shape is stable); the
  // dashed goal line is only *drawn* when a real target exists.
  const data: Datum[] = props.series.map((point, index) => ({
    x: index,
    y: point.value,
    target: target ?? point.value,
  }));

  return (
    <View style={{ height }}>
      <CartesianChart
        data={data}
        xKey="x"
        yKeys={['y', 'target']}
        domainPadding={{ top: 16, bottom: 16, left: 4, right: 4 }}
      >
        {({ points }) => (
          <>
            {target !== null && points.target ? (
              <Line points={points.target} color={theme.color.chartMuted} strokeWidth={1}>
                <DashPathEffect intervals={[6, 5]} />
              </Line>
            ) : null}
            <Line
              points={points.y}
              color={theme.color.chartLine}
              strokeWidth={2.5}
              curveType="natural"
            />
          </>
        )}
      </CartesianChart>
    </View>
  );
}
