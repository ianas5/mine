import { Dumbbell } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Card, EmptyState, ListRow, Section, Skeleton } from '@/core/ui';

/** Structure primitives: Card variants, Section, ListRow, EmptyState, Skeleton. */
export function GalleryStructure(): ReactNode {
  const theme = useTheme();
  const bodyText = { ...theme.type.body, color: theme.color.textPrimary } as const;
  return (
    <>
      <Section title="Card" action={{ label: 'See all', onPress: () => undefined }}>
        <View style={{ gap: theme.space.md }}>
          <Card>
            <Text style={bodyText}>default</Text>
          </Card>
          <Card variant="raised">
            <Text style={bodyText}>raised</Text>
          </Card>
          <Card variant="accentEdge">
            <Text style={bodyText}>accentEdge (InsightCard base)</Text>
          </Card>
        </View>
      </Section>
      <Section title="ListRow">
        <Card>
          <ListRow
            title="Bench Press"
            subtitle="Chest · last 4d ago"
            leading={<Dumbbell color={theme.color.accent} size={20} strokeWidth={1.75} />}
            trailingValue="85 kg"
            chevron
            onPress={() => undefined}
          />
          <ListRow title="Plain row" trailingValue="12" />
          <ListRow title="Title only" />
        </Card>
      </Section>
      <Section title="EmptyState">
        <Card>
          <EmptyState
            icon={<Dumbbell color={theme.color.textTertiary} size={28} strokeWidth={1.75} />}
            title="No weigh-ins yet"
            cta={{ label: 'Add Weight', onPress: () => undefined }}
          />
        </Card>
      </Section>
      <Section title="Skeleton">
        <View style={{ gap: theme.space.sm }}>
          <Skeleton />
          <Skeleton width="60%" />
          <Skeleton width={120} height={32} radius={theme.radius.md} />
        </View>
      </Section>
    </>
  );
}
