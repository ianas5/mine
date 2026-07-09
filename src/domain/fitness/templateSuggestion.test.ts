import { suggestTemplate, type RecentTemplateUse } from './templateSuggestion';

const use = (weekday: number, templateId: string): RecentTemplateUse => ({ weekday, templateId });

describe('suggestTemplate (UI_UX §5.2 chain)', () => {
  it('prefers the active program template scheduled for today', () => {
    expect(suggestTemplate(2, 'tpl_push', [use(2, 'tpl_other')], true)).toEqual({
      kind: 'template',
      templateId: 'tpl_push',
    });
  });

  it('falls back to the most-frequent template used on this weekday', () => {
    const recent = [use(1, 'tpl_a'), use(1, 'tpl_b'), use(1, 'tpl_b'), use(3, 'tpl_c')];
    expect(suggestTemplate(1, null, recent, true)).toEqual({
      kind: 'template',
      templateId: 'tpl_b',
    });
  });

  it('breaks a weekday-frequency tie toward the most recent use', () => {
    // newest-first: tpl_x is more recent than tpl_y, both used once on weekday 4.
    const recent = [use(4, 'tpl_x'), use(4, 'tpl_y')];
    expect(suggestTemplate(4, null, recent, true)).toEqual({
      kind: 'template',
      templateId: 'tpl_x',
    });
  });

  it('falls back to Repeat Last when nothing is scheduled or historical for today', () => {
    expect(suggestTemplate(5, null, [use(2, 'tpl_a')], true)).toEqual({ kind: 'repeatLast' });
  });

  it('suggests nothing when there is no plan and no history', () => {
    expect(suggestTemplate(5, null, [], false)).toEqual({ kind: 'none' });
  });
});
