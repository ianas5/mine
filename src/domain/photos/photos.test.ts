import { groupPhotosByDate, oldestMissingAngle, type PhotoAngle } from './index';

const cap = (angle: PhotoAngle, date: string) => ({ angle, date });

describe('oldestMissingAngle (UI_UX §5.2)', () => {
  it('prefers a never-captured angle, in front → side → back order', () => {
    expect(oldestMissingAngle([])).toBe('front');
    expect(oldestMissingAngle([cap('front', '2026-07-01')])).toBe('side');
    expect(oldestMissingAngle([cap('front', '2026-07-01'), cap('side', '2026-07-01')])).toBe(
      'back',
    );
  });

  it('when all captured, returns the angle whose most recent photo is oldest', () => {
    const photos = [
      cap('front', '2026-07-09'),
      cap('side', '2026-06-01'), // oldest most-recent
      cap('back', '2026-07-05'),
    ];
    expect(oldestMissingAngle(photos)).toBe('side');
  });
});

describe('groupPhotosByDate', () => {
  it('groups by date, newest first', () => {
    const groups = groupPhotosByDate([
      { date: '2026-07-01', angle: 'front' },
      { date: '2026-07-09', angle: 'front' },
      { date: '2026-07-01', angle: 'side' },
    ]);
    expect(groups.map((g) => g.date)).toEqual(['2026-07-09', '2026-07-01']);
    expect(groups[1]?.photos).toHaveLength(2);
  });
});
