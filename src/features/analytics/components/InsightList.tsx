import { useRouter, type Href } from 'expo-router';
import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Card, InsightCard } from '@/core/ui';
import { insightEvidenceHref, useInsights } from '@/data/analytics/useInsights';

/** The full insight list for the Analytics tab (UI_UX §8): every live insight, tappable to
 * its evidence and dismissible. A calm quiet state when nothing fires (§6.3). */
export function InsightList(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const insights = useInsights();

  if (insights === undefined) return null;

  if (insights.all.length === 0) {
    return (
      <Card>
        <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>No new signals</Text>
        <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
          Keep logging — insights appear as your data builds up.
        </Text>
      </Card>
    );
  }

  return (
    <View style={{ gap: theme.space.sm }}>
      {insights.all.map((insight) => (
        <InsightCard
          key={insight.instanceKey}
          tone={insight.tone}
          title={insight.title}
          body={insight.body}
          onPress={() => router.push(insightEvidenceHref(insight.evidence) as Href)}
          onDismiss={() => insights.dismiss(insight.instanceKey, insight.classification)}
        />
      ))}
    </View>
  );
}
