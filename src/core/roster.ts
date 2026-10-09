import { nameKey, normalize } from './normalize.ts';
import { SEED_PRESETS } from './seedPresets.ts';
import type { Rank, Session } from './types.ts';

export interface Staff {
  id: string;
  /** Name as printed on the sheet. */
  name: string;
  /** Spelling variants found in the masters, without the title. */
  aliases: string[];
  /** The `القسم` line of the sheet. */
  programme: string;
  rank: Rank;
  advising: boolean;
  /** 0 = Saturday … 5 = Thursday, or null. */
  secondmentDay: number | null;
  qualityHours: number;
  /** Office / advising hours to start from (default 4 each); both grow until the minimum load is met. */
  officeHours?: number;
  advisingHours?: number;
  /** Blocks this person keeps from term to term (see Preset). */
  presets?: Preset[];
}

/**
 * A remembered block of a person's week: where the committee put their office
 * hours / advising / quality, or a lecture or lab from a programme whose master
 * cannot be read yet (PDF only). Placed only where the masters leave the slot free.
 */
export interface Preset {
  day: number;
  slots: number[];
  type: 'lecture' | 'supervision' | 'section' | 'office' | 'advising' | 'quality';
  text: string;
  /** Row for the course table, for lectures from an unreadable master. */
  course?: { key: string; name: string; share: number; weeks: string };
}

const COMMS = 'هندسة الاتصالات والالكترونيات';
const POWER = 'هندسة القوى الكهربية';

const BASE: Staff[] = [
  { id: 'asmaa-radi', name: 'د/أسماء راضي', aliases: ['اسماء راضى', 'اسماء راضي'], programme: COMMS, rank: 'lecturer', advising: true, secondmentDay: 0, qualityHours: 2 },
  { id: 'ahmed-salem', name: 'د/أحمد ابراهيم سالم', aliases: ['احمد سالم', 'احمد ابراهيم سالم'], programme: COMMS, rank: 'lecturer', advising: true, secondmentDay: null, qualityHours: 2 },
  { id: 'mai-helmy', name: 'د/مي حلمي', aliases: ['مى حلمى', 'مي حلمي'], programme: COMMS, rank: 'lecturer', advising: true, secondmentDay: 4, qualityHours: 2 },
  { id: 'hend-elsayed', name: 'د/هند علي السيد', aliases: ['هند السيد', 'هند على السيد'], programme: COMMS, rank: 'associate', advising: true, secondmentDay: 1, qualityHours: 2 },
  { id: 'ahmed-elsayed', name: 'د/أحمد السيد متولي', aliases: ['احمد السيد', 'اخمد السيد', 'احمد السيد متولى'], programme: 'هندسة القوي الكهربية', rank: 'lecturer', advising: true, secondmentDay: 4, qualityHours: 2 },
  { id: 'asmaa-abdelrahim', name: 'د/ أسماء عبدالرحيم إبراهيم السقعان', aliases: ['اسماء عبدالرحيم', 'اسماء عبد الرحيم السقعان'], programme: POWER, rank: 'lecturer', advising: true, secondmentDay: 4, qualityHours: 2 },
  { id: 'mahmoud-elsadd', name: 'ا.د/محمود عبدآمين السد', aliases: ['محمود السد', 'محمود عبدامين السد'], programme: POWER, rank: 'professor', advising: false, secondmentDay: 2, qualityHours: 2 },
  { id: 'gomaa-fahmy', name: 'ا.م.د/جمعه فهمى عبدالنبى', aliases: ['جمعه فهمى', 'جمعة فهمى', 'جمعه فهمى عبدالنبى'], programme: POWER, rank: 'associate', advising: false, secondmentDay: null, qualityHours: 2 },
  { id: 'eman-awad', name: 'د/ايمان احمد عوض', aliases: ['ايمان عوض', 'ايمان احمد عوض'], programme: POWER, rank: 'lecturer', advising: true, secondmentDay: 5, qualityHours: 2 },
  { id: 'tamer-elsharkawy', name: 'م/تامر الشرقاوى', aliases: ['تامر الشرقاوى', 'تامر'], programme: POWER, rank: 'ta', advising: true, secondmentDay: null, qualityHours: 2 },
];

// Office / advising / quality hours as on the hand-made sheets (brief §10).
const HOURS: Record<string, [number, number, number]> = {
  'asmaa-radi': [4, 4, 2], 'ahmed-salem': [4, 0, 2], 'mai-helmy': [5, 4, 2], 'hend-elsayed': [4, 4, 2],
  'ahmed-elsayed': [4, 4, 2], 'asmaa-abdelrahim': [5, 6, 5], 'mahmoud-elsadd': [6, 0, 4], 'gomaa-fahmy': [8, 0, 4],
  'eman-awad': [6, 5, 2], 'tamer-elsharkawy': [0, 6, 5],
};

/** Seed roster from brief §6 and the hand-made sheets. Editable in the UI; this is only the starting point. */
export const SEED_ROSTER: Staff[] = BASE.map((staff) => {
  const [officeHours, advisingHours, qualityHours] = HOURS[staff.id] ?? [4, 4, 2];
  return { ...staff, advising: advisingHours > 0, officeHours, advisingHours, qualityHours, presets: SEED_PRESETS[staff.id] ?? [] };
});

/**
 * Fill `staffId` on every name token by exact alias match (letters only), then
 * scan the whole cell for full-name aliases the tokens missed because of broken
 * punctuation. One-word names (`د. اسماء`) are never guessed here — they are
 * resolved from the course's lecture or on the review screen (brief §4).
 */
export function matchStaff(sessions: Session[], roster: Staff[]): void {
  const byKey = new Map<string, string>();
  const scans: { id: string; re: RegExp }[] = [];
  for (const staff of roster) {
    for (const alias of [staff.name.replace(/^[^/]*\//, ''), ...staff.aliases]) {
      const words = normalize(alias).split(' ').filter(Boolean);
      byKey.set(nameKey(alias), staff.id);
      if (words.length > 1) {
        scans.push({ id: staff.id, re: new RegExp(`(?<!\\p{L})${words.join('\\s*')}(?!\\p{L})`, 'u') });
      }
    }
  }
  for (const s of sessions) {
    const found = new Set<string>();
    for (const token of s.parsed.names) {
      const isTa = roster.find((x) => x.id === byKey.get(nameKey(token.name)))?.rank === 'ta';
      // A TA is never mistaken for a doctor of the same name, and vice versa.
      token.staffId = (token.role === 'ta') === isTa ? (byKey.get(nameKey(token.name)) ?? null) : null;
      if (token.staffId) found.add(token.staffId);
    }
    const text = normalize(s.text);
    s.scannedStaff = [...new Set(scans.filter((x) => !found.has(x.id) && x.re.test(text)).map((x) => x.id))];
  }
}
