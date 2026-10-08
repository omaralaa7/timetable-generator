import { nameKey, normalize } from './normalize.ts';
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
}

const COMMS = 'هندسة الاتصالات والالكترونيات';
const POWER = 'هندسة القوى الكهربية';

/** Seed roster from brief §6. Editable in the UI; this is only the starting point. */
export const SEED_ROSTER: Staff[] = [
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
