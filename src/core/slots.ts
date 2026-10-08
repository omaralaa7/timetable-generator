import { normalize } from './normalize.ts';

export const DAYS = ['السبت', 'الاحد', 'الاثنين', 'الثلاثاء', 'الاربعاء', 'الخميس'];

/** The 10 fixed slots (brief §3.3) as [start, end] in minutes since midnight. */
export const SLOTS: [number, number][] = [
  [540, 590], [590, 640], [645, 695], [695, 745], [750, 800],
  [800, 850], [855, 905], [905, 955], [960, 1010], [1010, 1060],
];

export const SLOT_LABELS = [
  '09:00-09:50', '09:50-10:40', '10:45-11:35', '11:35-12:25', '12:30-01:20',
  '01:20-02:10', '02:15-03:05', '03:05-03:55', '04:00-04:50', '04:50-05:40',
];

const DAY_KEYS = DAYS.map((d) => normalize(d));

export function parseDay(text: string): number | null {
  // Repeated letters are typos (`االثلاثاء`).
  const t = normalize(text).replace(/\s/g, '').replace(/(.)\1+/gu, '$1');
  if (!t) return null;
  const i = DAY_KEYS.indexOf(t);
  if (i >= 0) return i;
  // `الأحد` / `الإثنين` / `الأربعاء` already collapse to the keys above; allow the article to be missing.
  const j = DAY_KEYS.findIndex((d) => d.slice(2) === t);
  return j >= 0 ? j : null;
}

function times(label: string): number[] {
  const out: number[] = [];
  for (const m of normalize(label).matchAll(/(\d{1,2})\s*[:.]\s*(\d{2})/g)) {
    let h = Number(m[1]);
    if (h < 8) h += 12; // the timetable runs 09:00 → 05:40 pm
    out.push(h * 60 + Number(m[2]));
  }
  return out;
}

function closest(start: number, end: number | undefined): { slot: number; off: number } {
  let best = { slot: -1, off: Infinity };
  SLOTS.forEach(([s, e], slot) => {
    const off = Math.abs(start - s) + (end === undefined ? 0 : Math.abs(end - e));
    if (off < best.off) best = { slot, off };
  });
  return best;
}

/**
 * Map a time label to a slot by its start time. Tolerates typos such as
 * `09.40`, `08:50`, a missing dash, or swapped start/end (`5:40 – 4:50`).
 * `exact` is false when the label had to be bent to fit.
 */
export function parseSlot(label: string): { slot: number; exact: boolean } | null {
  const t = times(label);
  if (!t.length) return null;
  const tries = [closest(t[0], t[1])];
  if (t.length > 1) tries.push(closest(t[1], t[0]), closest(t[0], undefined));
  const exact = tries[0].off === 0;
  const best = tries.reduce((a, b) => (b.off < a.off ? b : a));
  return best.off <= 30 ? { slot: best.slot, exact } : null;
}
