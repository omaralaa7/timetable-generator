import { nameKey, normalize } from './normalize.ts';
import type { Preset, Staff } from './roster.ts';
import { SEED_STARTS } from './seedStarts.ts';
import type { NameToken, ParsedText, Rank, Session } from './types.ts';

export type ItemType = 'lecture' | 'supervision' | 'section' | 'office' | 'advising' | 'quality';

/** One block in a person's weekly grid. */
export interface GridItem {
  day: number;
  slots: number[];
  type: ItemType;
  text: string;
  sessionId?: string;
  /** `auto` = suggested by the tool (duties), `master` = read from a master. */
  origin: 'master' | 'auto' | 'manual';
}

export interface CourseRow {
  key: string;
  name: string;
  share: number;
  /** `1 - 7`, `9 - 15`, `1 - 15` or `؟` while undecided. */
  weeks: string;
}

export type IssueKind =
  | 'no-underline'
  | 'both-free'
  | 'conflict'
  | 'overlap'
  | 'short-name'
  | 'below-minimum'
  | 'project';

export interface Issue {
  id: string;
  kind: IssueKind;
  message: string;
  staffIds: string[];
  sessionId?: string;
  courseKey?: string;
  /** Roster ids the user can pick from, the first being the tool's proposal. */
  options?: string[];
  /** Display name of each option. */
  labels?: Record<string, string>;
}

export interface Totals {
  /** Lecture hours — or section hours for a teaching assistant. */
  teaching: number;
  office: number;
  supervision: number;
  advising: number;
  quality: number;
  /** Everything except quality. */
  counted: number;
  minimum: number;
}

export interface StaffSheet {
  staff: Staff;
  courses: CourseRow[];
  items: GridItem[];
  daily: number[];
  totals: Totals;
}

/** Choices made on the review screen. Everything is optional; missing = let the tool decide. */
export interface Decisions {
  /** course key → roster id / name key of the doctor who teaches weeks 1–7. */
  starts?: Record<string, string>;
  /** session id → roster id of the supervising doctor, or `none`. */
  supervisor?: Record<string, string>;
  /** name key as written → roster id, or `ignore`. */
  names?: Record<string, string>;
  /** roster id → the duty blocks, replacing the tool's suggestion. */
  duties?: Record<string, GridItem[]>;
  /** roster id → `day:slot` → what the user put in that cell (null = empty). Applied last. */
  edits?: Record<string, Record<string, { type: ItemType; text: string } | null>>;
  /** roster id → the course table as edited by the user. */
  courses?: Record<string, CourseRow[]>;
  /** issue id → hidden by the user. */
  dismissed?: Record<string, boolean>;
}

export interface Result {
  sheets: StaffSheet[];
  issues: Issue[];
  /** Doctors/assistants named in the masters but absent from the roster. */
  detected: { name: string; role: NameToken['role']; count: number }[];
}

export const MINIMUM: Record<Rank, number> = { professor: 25, associate: 27, lecturer: 29, ta: 33 };
const DUTY_TEXT = { office: 'ساعات مكتبية', advising: 'ارشاد أكاديمي', quality: 'جودة' } as const;
const PAIRS = [[0, 1], [2, 3], [4, 5], [6, 7], [8, 9]];

/**
 * Who teaches weeks 1–7 where the master has no underline — answered by the client (2026-10-09).
 * The lasting fix is to underline the name in the master; an underline or a review-screen choice wins over this.
 */
const CLIENT_STARTS: { code: string; title?: string; subtitle?: string; first: string }[] = [
  { code: 'MAT 201', first: 'اسماء عبدالرحيم' },
  { code: 'ELP 3E1', subtitle: 'نظم القياسات', first: 'سيد محمد احمد' },
  // Preparatory level, as on the hand-made sheets (the master has no underline there).
  { code: 'MAT 001', first: 'جمعة فهمى' },
  { code: 'HUM 002', first: 'اسماء عبدالرحيم' },
  // Mechanics, as on the hand-made sheets. ELP 331 is a different course in the power programme.
  { code: 'ELP 331', title: 'الات القوى الكهربية', first: 'احمد السيد' },
];

/**
 * A doctor who teaches a lecture but is not written in the master (brief §9): the mechanics
 * timetable lists only Dr. Mai Helmy for this course; Dr. Asmaa Radi teaches weeks 1–7.
 */
const CLIENT_COTEACHERS: { title: string; name: string; starts: boolean }[] = [
  { title: 'دوائر كهربية والكترونية', name: 'اسماء راضى', starts: true },
];

const key = (s: string) => normalize(s).replace(/[^\p{L}\p{N}]/gu, '');
const isElectiveKind = (p: ParsedText) => /اختياري/.test(normalize(p.kind));

export function courseKey(p: ParsedText): string {
  return [p.code.replace(/\s/g, ''), key(p.title), key(p.subtitle)].join('|');
}

export function courseName(p: ParsedText, withCode = true): string {
  return [
    isElectiveKind(p) ? p.kind : '',
    p.title,
    withCode && p.code ? `(${p.code})` : '',
    p.subtitle ? `(${p.subtitle})` : '',
  ].filter(Boolean).join(' ');
}

export function lectureText(p: ParsedText): string {
  const kind = p.kind || 'محاضرة';
  return [kind, p.title, p.code ? `(${p.code})` : '', p.subtitle ? `(${p.subtitle})` : '', p.room ? `/ ${p.room}` : '']
    .filter(Boolean).join(' ');
}

export function sectionText(p: ParsedText, supervised: boolean): string {
  return [supervised ? 'اشراف' : '', p.kind, p.title, p.room ? `/ ${p.room}` : ''].filter(Boolean).join(' ');
}

const overlaps = (a: { day: number; slots: number[] }, b: { day: number; slots: number[] }) =>
  a.day === b.day && a.slots.some((s) => b.slots.includes(s));

/**
 * Apply the client's rules (brief §5) to the sessions of every uploaded master
 * and build each roster member's sheet. Pure: same input → same output.
 */
export function buildSheets(sessions: Session[], roster: Staff[], decisions: Decisions = {}): Result {
  const issues: Issue[] = [];
  const byId = new Map(roster.map((s) => [s.id, s]));
  const items = new Map<string, GridItem[]>(roster.map((s) => [s.id, []]));
  const courses = new Map<string, Map<string, CourseRow & { order: number }>>(roster.map((s) => [s.id, new Map()]));
  const busy = (id: string, at: { day: number; slots: number[] }) =>
    byId.get(id)!.secondmentDay === at.day || items.get(id)!.some((it) => overlaps(it, at)) ||
    at.slots.some((s) => decisions.edits?.[id]?.[`${at.day}:${s}`]); // a cell the user filled by hand

  // Weeks of each placed lecture, to tell a real clash from two half-term lectures sharing a slot.
  const lectureWeeks = new Map<GridItem, string>();
  const place = (id: string, item: GridItem, what: string, weeks = '') => {
    const staff = byId.get(id)!;
    const clash = items.get(id)!.find((it) => overlaps(it, item));
    if (weeks) lectureWeeks.set(item, weeks);
    const halves = [weeks, clash ? lectureWeeks.get(clash) : ''].sort().join('|');
    if (clash && halves === '1 - 7|9 - 15') {
      // One lecture in weeks 1–7 and another in weeks 9–15 at the same time is not a clash.
      // The hand-made sheets show a single lecture in that cell: the one placed first stays.
      return;
    }
    if (clash) {
      issues.push({
        id: `overlap:${id}:${item.sessionId}`, kind: 'overlap', staffIds: [id], sessionId: item.sessionId,
        message: `${staff.name}: «${what}» يتعارض مع «${clash.text}»`,
      });
    } else if (staff.secondmentDay === item.day) {
      issues.push({
        id: `overlap:${id}:${item.sessionId}`, kind: 'overlap', staffIds: [id], sessionId: item.sessionId,
        message: `${staff.name}: «${what}» يقع في يوم الانتداب`,
      });
    }
    items.get(id)!.push(item);
  };

  // ── Names: review-screen mappings, then one-word names from the course's lecture ──
  const lectures = sessions.filter((s) => s.parsed.type === 'lecture');
  for (const s of lectures) {
    for (const extra of CLIENT_COTEACHERS.filter((c) => key(c.title) === key(s.parsed.title))) {
      const doctors = s.parsed.names.filter((t) => t.role === 'doctor');
      const staff = roster.find((r) => r.aliases.some((a) => nameKey(a) === nameKey(extra.name)));
      if (!staff || doctors.length !== 1 || doctors[0].staffId === staff.id) continue;
      if (extra.starts) doctors[0].underline = 0;
      s.parsed.names.push({ name: extra.name, prefix: 'د', role: 'doctor', rank: 'lecturer', underline: extra.starts ? 1 : 0, staffId: staff.id });
    }
  }
  const detected = new Map<string, { name: string; role: NameToken['role']; count: number }>();
  for (const s of sessions) {
    for (const t of s.parsed.names) {
      const k = nameKey(t.name);
      const mapped = decisions.names?.[k];
      if (!t.staffId && mapped && byId.has(mapped)) t.staffId = mapped;
      if (t.staffId || mapped === 'ignore') continue;
      const oneWord = !normalize(t.name).includes(' ');
      if (oneWord && t.role === 'doctor') {
        const pool = lectures
          .filter((l) => l.programme === s.programme && (key(l.parsed.title) === key(s.parsed.title) ||
            (l.parsed.subtitle && key(s.parsed.title).includes(key(l.parsed.subtitle))) ||
            // `م برمجة` is the lab of `حاسبات وبرمجة`
            (key(s.parsed.title).length >= 4 && key(l.parsed.title).includes(key(s.parsed.title)))))
          .flatMap((l) => l.parsed.names)
          .filter((n) => n.role === 'doctor' && normalize(n.name).split(' ')[0] === normalize(t.name));
        const ids = [...new Set(pool.map((n) => n.staffId ?? `?${nameKey(n.name)}`))];
        if (ids.length === 1) {
          if (!ids[0].startsWith('?')) t.staffId = ids[0];
          continue; // resolved, possibly to someone outside the department
        }
        const options = roster.filter((r) => r.aliases.some((a) => normalize(a).split(' ')[0] === normalize(t.name))).map((r) => r.id);
        if (options.length && !issues.some((i) => i.id === `short:${k}`)) {
          issues.push({
            id: `short:${k}`, kind: 'short-name', staffIds: options, options,
            message: `الاسم المختصر «${t.prefix} ${t.name}» في «${s.text}» — من المقصود؟`,
          });
        }
        continue;
      }
      const d = detected.get(`${t.role}:${k}`) ?? { name: `${t.prefix} ${t.name}`.trim(), role: t.role, count: 0 };
      d.count++;
      detected.set(`${t.role}:${k}`, d);
    }
  }
  const rosterDoctors = (s: Session) => {
    const ids = s.parsed.names.filter((t) => t.role === 'doctor' && t.staffId).map((t) => t.staffId!);
    for (const id of s.scannedStaff) if (byId.get(id)?.rank !== 'ta') ids.push(id);
    return [...new Set(ids)];
  };

  // ── Lectures (§5.1, §5.2). A level-1 lecture shared by two programmes is one lecture. ──
  // Same time, a doctor in common and the same room or course → the same lecture; the fuller cell is kept.
  const doctorKeys = (s: Session) => s.parsed.names.filter((t) => t.role === 'doctor').map((t) => nameKey(t.name));
  const sameLecture = (a: Session, b: Session) =>
    a.day === b.day && a.slots.join() === b.slots.join() && doctorKeys(a).some((k) => doctorKeys(b).includes(k)) &&
    // the same room, or the same course with the room written differently (`مدرج 7` / `مدرج كهرباء`)
    (key(a.parsed.room) === key(b.parsed.room) || courseKey(a.parsed) === courseKey(b.parsed));
  const unique: Session[] = [];
  for (const s of [...lectures].sort((a, b) => doctorKeys(b).length - doctorKeys(a).length)) {
    if (!unique.some((u) => sameLecture(u, s))) unique.push(s);
  }
  // A course taught in several programmes: one underlined cell settles the weeks for all of them.
  const idOf = (t: NameToken) => t.staffId ?? nameKey(t.name);
  const startOf = new Map<string, string>();
  for (const s of unique) {
    const doctors = s.parsed.names.filter((t) => t.role === 'doctor');
    const marked = doctors.filter((t) => t.underline > 0.5);
    if (doctors.length === 2 && marked.length === 1) startOf.set(courseKey(s.parsed), idOf(marked[0]));
  }
  let order = 0;
  for (const s of unique.sort((a, b) => (a.level ?? 9) - (b.level ?? 9) || a.day - b.day || a.slots[0] - b.slots[0])) {
    const p = s.parsed;
    const doctors = p.names.filter((t) => t.role === 'doctor');
    const mine = rosterDoctors(s);
    if (!mine.length) continue;
    const ck = courseKey(p);
    let first: string | null = null; // who teaches weeks 1–7
    if (doctors.length === 2) {
      const marked = doctors.filter((t) => t.underline > 0.5);
      const chosen = decisions.starts?.[ck];
      if (chosen && doctors.some((t) => idOf(t) === chosen)) first = chosen;
      else if (marked.length === 1) first = idOf(marked[0]);
      else if (doctors.some((t) => idOf(t) === startOf.get(ck))) first = startOf.get(ck)!;
      else {
        const told = CLIENT_STARTS.find((c) => c.code === p.code && (!c.subtitle || key(c.subtitle) === key(p.subtitle)) &&
          (!c.title || key(c.title) === key(p.title)));
        const starter = told && doctors.find((t) => nameKey(t.name) === nameKey(told.first));
        if (starter) first = idOf(starter);
        // Last resort: what an earlier, underlined copy of the masters said for this course.
        else if (doctors.some((t) => idOf(t) === SEED_STARTS[ck])) first = SEED_STARTS[ck];
      }
      if (!first && !issues.some((i) => i.id === `start:${ck}`)) {
        issues.push({
          id: `start:${ck}`, kind: 'no-underline', staffIds: mine, courseKey: ck, options: doctors.map(idOf),
          labels: Object.fromEntries(doctors.map((t) => [idOf(t), `${t.prefix} ${t.name}`.trim()])),
          message: `«${courseName(p)}»: دكتوران بدون خط تحت أحدهما — من يبدأ (الأسابيع 1–7)؟`,
        });
      }
    } else if (doctors.length > 2 && !issues.some((i) => i.id === `start:${ck}`)) {
      issues.push({
        id: `start:${ck}`, kind: 'no-underline', staffIds: mine, courseKey: ck,
        message: `«${courseName(p)}»: أكثر من دكتورين في المحاضرة — راجع النسبة والأسابيع`,
      });
    }
    for (const id of mine) {
      const share = doctors.length <= 1 ? 1 : Math.round(100 / doctors.length) / 100;
      const weeks = doctors.length <= 1 ? '1 - 15' : doctors.length > 2 || !first ? '؟' : first === id ? '1 - 7' : '9 - 15';
      if (!courses.get(id)!.has(ck)) courses.get(id)!.set(ck, { key: ck, name: courseName(p), share, weeks, order: order++ });
      place(id, { day: s.day, slots: s.slots, type: 'lecture', text: lectureText(p), sessionId: s.id, origin: 'master' }, lectureText(p), weeks);
    }
  }

  // Lectures and labs remembered from programmes whose master cannot be read yet.
  const fits = (id: string, p: Preset) => byId.get(id)!.secondmentDay !== p.day && !items.get(id)!.some((it) => overlaps(it, p));
  for (const staff of roster) {
    for (const p of staff.presets ?? []) {
      if (p.type !== 'lecture' && p.type !== 'supervision' && p.type !== 'section') continue;
      // The master of that programme has been uploaded since: it wins and the remembered copy is dropped.
      const name = p.course ? key(p.course.name.replace(/\(.*$/, '')) : '';
      if (name && [...courses.get(staff.id)!.values()].some((c) => !c.key.startsWith('external|') && key(c.name).includes(name))) continue;
      if (p.course && !courses.get(staff.id)!.has(p.course.key)) courses.get(staff.id)!.set(p.course.key, { ...p.course, order: order++ });
      // Once the user has edited this person's grid, their blocks live in `decisions.duties`.
      if (decisions.duties?.[staff.id] || !fits(staff.id, p)) continue;
      items.get(staff.id)!.push({ day: p.day, slots: [...p.slots], type: p.type, text: p.text, origin: 'manual' });
    }
  }

  // ── Sections and labs (§5.3) ──
  const sections = sessions
    .filter((s) => s.parsed.type === 'section')
    .sort((a, b) => a.day - b.day || a.slots[0] - b.slots[0] || a.id.localeCompare(b.id));
  for (const s of sections) {
    // A teaching assistant takes every session they are named in.
    for (const t of s.parsed.names) {
      if (t.role !== 'ta' || !t.staffId) continue;
      const text = sectionText(s.parsed, false);
      const ck = `|${key(s.parsed.title)}|`;
      if (!courses.get(t.staffId)!.has(ck)) {
        courses.get(t.staffId)!.set(ck, { key: ck, name: courseName(s.parsed, false), share: 1, weeks: '1 - 15', order: order++ });
      }
      place(t.staffId, { day: s.day, slots: s.slots, type: 'section', text, sessionId: s.id, origin: 'master' }, text);
    }
  }
  const supervisor = new Map<string, string>(); // session id → roster id
  const sectionOf = (s: Session) => `${s.programme}|${s.level}|${key(s.parsed.title)}`;
  const assign = (s: Session, id: string) => {
    supervisor.set(s.id, id);
    const text = sectionText(s.parsed, true);
    items.get(id)!.push({ day: s.day, slots: s.slots, type: 'supervision', text, sessionId: s.id, origin: 'master' });
  };
  // How far a person is from their minimum once their usual office hours and advising are added (negative = still short).
  const load = (id: string) => {
    const staff = byId.get(id)!;
    const duties = (staff.officeHours ?? 4) + (staff.advising ? (staff.advisingHours ?? 4) : 0);
    return items.get(id)!.reduce((n, it) => n + it.slots.length, 0) + duties - MINIMUM[staff.rank];
  };
  let pending = sections.filter((s) => rosterDoctors(s).length > 0);
  for (const s of pending) {
    const chosen = decisions.supervisor?.[s.id];
    if (chosen === 'none') supervisor.set(s.id, 'none');
    else if (chosen && byId.has(chosen)) assign(s, chosen);
  }
  pending = pending.filter((s) => !supervisor.has(s.id));
  while (pending.length) {
    // Forced: exactly one of the listed doctors is free.
    const forced = pending.find((s) => rosterDoctors(s).filter((id) => !busy(id, s)).length === 1);
    if (forced) {
      assign(forced, rosterDoctors(forced).find((id) => !busy(id, forced))!);
      pending = pending.filter((s) => s !== forced);
      continue;
    }
    const open = pending.find((s) => rosterDoctors(s).some((id) => !busy(id, s)));
    if (!open) break;
    // Both free: the tool proposes one and the user can swap on the review screen.
    const free = rosterDoctors(open).filter((id) => !busy(id, open));
    const siblings = sections.filter((x) => supervisor.has(x.id) && sectionOf(x) === sectionOf(open));
    const sameSection = siblings.find((x) => x.sections.join() === open.sections.join() && free.includes(supervisor.get(x.id)!));
    const otherSection = siblings.find((x) => x.sections.join() !== open.sections.join());
    // Narrow the choice step by step, the way the committee does it by hand (client, 2026-10-09):
    // lectures are placed first, then supervision; office hours and advising come last and never decide this.
    const courseHours = (id: string) => sections
      .filter((x) => supervisor.get(x.id) === id && key(x.parsed.title) === key(open.parsed.title))
      .reduce((n, x) => n + x.slots.length, 0);
    // In the faculty that day: something already placed, or a day this person is known to attend.
    const present = (id: string) => items.get(id)!.some((it) => it.day === open.day) ||
      (byId.get(id)!.presets ?? []).some((p) => p.day === open.day);
    let pool = free;
    const prefer = (test: (id: string) => boolean) => {
      const kept = pool.filter(test);
      if (kept.length) pool = kept;
    };
    prefer(present); // the doctor who is already in the faculty that day
    // Section 1 goes to one doctor and section 2 to the other; a single section is never split between them.
    if (sameSection) prefer((id) => id === supervisor.get(sameSection.id));
    prefer((id) => courseHours(id) === Math.min(...pool.map(courseHours))); // the two doctors share a course's sessions evenly
    if (otherSection) prefer((id) => id !== supervisor.get(otherSection.id));
    // Otherwise it goes to whoever still needs the hours to complete their minimum load.
    const pick = [...pool].sort((x, y) => load(x) - load(y))[0];
    assign(open, pick);
    issues.push({
      id: `free:${open.id}`, kind: 'both-free', staffIds: free, sessionId: open.id, options: [pick, ...free.filter((id) => id !== pick)],
      message: `«${open.text}»: الدكتوران متاحان — المقترح ${byId.get(pick)!.name}`,
    });
    pending = pending.filter((s) => s !== open);
  }
  // Sessions left in `pending` have no listed doctor free. The committee's rule (client, 2026-10-09):
  // a doctor supervises only where free and in the faculty, so these simply get no supervisor
  // from the department — it is not a problem to resolve, and nothing is reported.
  for (const s of sessions.filter((x) => x.parsed.type === 'project' && rosterDoctors(x).length)) {
    issues.push({ id: `project:${s.id}`, kind: 'project', staffIds: rosterDoctors(s), sessionId: s.id, message: `مشروع التخرج يحتاج مراجعة: «${s.text}»` });
  }

  // ── Duties (§5.4) and totals ──
  const sheets = roster.map((staff): StaffSheet => {
    const mine = items.get(staff.id)!;
    const manual = decisions.duties?.[staff.id];
    if (manual) mine.push(...manual.map((d) => ({ ...d, origin: 'manual' as const })));
    else suggestDuties(staff, mine);
    applyEdits(mine, decisions.edits?.[staff.id]);
    mine.sort((a, b) => a.day - b.day || a.slots[0] - b.slots[0]);
    // The same lecture given to two groups back to back is one block on the sheet.
    for (let i = mine.length - 1; i > 0; i--) {
      const prev = mine[i - 1];
      const it = mine[i];
      if (prev.day === it.day && prev.type === it.type && prev.text === it.text && prev.slots[prev.slots.length - 1] + 1 === it.slots[0]) {
        prev.slots = [...prev.slots, ...it.slots];
        mine.splice(i, 1);
      }
    }
    const hours = (type: ItemType) => mine.filter((it) => it.type === type).reduce((n, it) => n + it.slots.length, 0);
    const isTa = staff.rank === 'ta';
    const totals: Totals = {
      teaching: hours(isTa ? 'section' : 'lecture'),
      office: hours('office'),
      supervision: hours('supervision'),
      advising: hours('advising'),
      quality: hours('quality'),
      counted: 0,
      minimum: MINIMUM[staff.rank],
    };
    totals.counted = totals.teaching + totals.office + totals.supervision + totals.advising;
    if (totals.counted < totals.minimum) {
      issues.push({
        id: `min:${staff.id}`, kind: 'below-minimum', staffIds: [staff.id],
        message: `${staff.name}: النصاب ${totals.counted} من ${totals.minimum}`,
      });
    }
    const daily = [0, 1, 2, 3, 4, 5].map((d) => new Set(mine.filter((it) => it.day === d).flatMap((it) => it.slots)).size);
    const rows = decisions.courses?.[staff.id] ??
      [...courses.get(staff.id)!.values()].sort((a, b) => a.order - b.order).map(({ key, name, share, weeks }) => ({ key, name, share, weeks }));
    return { staff, courses: rows, items: mine, daily, totals };
  });

  for (const i of issues) {
    if (i.options && !i.labels) i.labels = Object.fromEntries(i.options.map((o) => [o, byId.get(o)?.name ?? o]));
  }
  return { sheets, issues: issues.filter((i) => !decisions.dismissed?.[i.id]), detected: [...detected.values()].sort((a, b) => b.count - a.count) };
}

/**
 * Suggest quality, office-hour and advising blocks in free slots: 2-slot blocks
 * first, never on the secondment day, Thursday only when nothing else is left.
 * Office hours and advising grow until the minimum load is reached.
 */
function suggestDuties(staff: Staff, items: GridItem[]): void {
  const taken = (day: number, slot: number) => items.some((it) => it.day === day && it.slots.includes(slot));
  const dayLoad = (day: number) => items.filter((it) => it.day === day).reduce((n, it) => n + it.slots.length, 0);
  const add = (type: 'office' | 'advising' | 'quality', want: number): number => {
    let left = want;
    for (const size of [2, 1]) {
      while (left >= size) {
        const spots: { day: number; slots: number[]; rank: number }[] = [];
        for (let day = 0; day < 6; day++) {
          if (day === staff.secondmentDay) continue;
          const blocks = size === 2 ? PAIRS : PAIRS.flat().map((s) => [s]);
          for (const slots of blocks) {
            if (slots.some((s) => taken(day, s))) continue;
            // Thursday last, then days the person already comes in, then early in the day.
            spots.push({ day, slots, rank: (day === 5 ? 1000 : 0) + (dayLoad(day) ? 0 : 100) + slots[0] * 2 + day });
          }
        }
        if (!spots.length) break;
        const best = spots.sort((a, b) => a.rank - b.rank)[0];
        items.push({ day: best.day, slots: best.slots, type, text: DUTY_TEXT[type], origin: 'auto' });
        left -= size;
      }
    }
    return want - left;
  };
  const hours = (type: ItemType) => items.filter((it) => it.type === type).reduce((n, it) => n + it.slots.length, 0);
  const counted = () => items.filter((it) => it.type !== 'quality').reduce((n, it) => n + it.slots.length, 0);
  const canAdvise = staff.advising;
  // First the places the committee already chose, wherever they are still free …
  for (const p of staff.presets ?? []) {
    if (p.type !== 'office' && p.type !== 'advising' && p.type !== 'quality') continue;
    if (p.day === staff.secondmentDay || (p.type === 'advising' && !canAdvise)) continue;
    const want = { office: staff.officeHours ?? 4, advising: staff.advisingHours ?? 4, quality: staff.qualityHours }[p.type];
    for (const run of p.slots.filter((s) => !taken(p.day, s)).map((s) => [s])) {
      if (hours(p.type) >= want) break;
      const last = items[items.length - 1];
      if (last && last.origin === 'auto' && last.type === p.type && last.day === p.day && last.slots[last.slots.length - 1] === run[0] - 1) last.slots.push(run[0]);
      else items.push({ day: p.day, slots: run, type: p.type, text: DUTY_TEXT[p.type], origin: 'auto' });
    }
  }
  // … then whatever is still missing.
  if (canAdvise) add('advising', Math.max(0, (staff.advisingHours ?? 4) - hours('advising')));
  add('office', Math.max(0, (staff.officeHours ?? 4) - hours('office')));
  // Top up to the minimum, alternating office hours and advising.
  let turn: 'office' | 'advising' = 'office';
  while (counted() < MINIMUM[staff.rank]) {
    const need = Math.min(2, MINIMUM[staff.rank] - counted());
    if (!add(canAdvise ? turn : 'office', need)) break;
    turn = turn === 'office' ? 'advising' : 'office';
  }
  add('quality', Math.max(0, staff.qualityHours - hours('quality')));
}

/** Cell-by-cell changes from the editable grid win over everything else. */
function applyEdits(items: GridItem[], edits: Record<string, { type: ItemType; text: string } | null> = {}): void {
  for (const [at, value] of Object.entries(edits).sort(([a], [b]) => a.localeCompare(b, 'en', { numeric: true }))) {
    const [day, slot] = at.split(':').map(Number);
    for (const it of [...items]) {
      if (it.day !== day || !it.slots.includes(slot)) continue;
      const after = it.slots.filter((x) => x > slot);
      it.slots = it.slots.filter((x) => x < slot);
      // The cell may be cut out of the middle of a block: keep both halves.
      if (after.length) items.push({ ...it, slots: after });
      if (!it.slots.length) items.splice(items.indexOf(it), 1);
    }
    if (!value) continue;
    const before = items.find((it) => it.day === day && it.type === value.type && it.text === value.text &&
      it.origin === 'manual' && it.slots[it.slots.length - 1] === slot - 1);
    if (before) before.slots.push(slot);
    else items.push({ day, slots: [slot], type: value.type, text: value.text, origin: 'manual' });
  }
}
