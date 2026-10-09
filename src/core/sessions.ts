import { parseCellText } from './cellText.ts';
import { normalize, tidy } from './normalize.ts';
import { parseDay, parseSlot } from './slots.ts';
import type { ParsedMaster, RawCell, RawTable, Session, Warning } from './types.ts';

interface SectionCol {
  start: number;
  end: number; // exclusive
  level: number | null;
  section: number;
  width: number;
}

const ORDINALS = ['الاول', 'الثاني', 'الثالث', 'الرابع', 'الخامس'];

function cellText(cell: RawCell): string {
  return cell.runs.map((r) => r.text).join('');
}

function levelNumber(text: string): number | null {
  const n = normalize(text);
  const m = /(?:مستوي|فرقه)\s*(\d+)/.exec(n);
  if (m) return Number(m[1]);
  const i = ORDINALS.findIndex((o) => n.includes(o));
  return i >= 0 && /مستوي|فرقه/.test(n) ? i + 1 : null;
}

/** Column with the most cells accepted by `test` (the day column, the time column). */
function bestColumn(table: RawTable, test: (text: string) => boolean): number {
  const hits = new Map<number, number>();
  for (const c of table.cells) if (test(cellText(c))) hits.set(c.col, (hits.get(c.col) ?? 0) + 1);
  let best = -1;
  for (const [col, n] of hits) if (best < 0 || n > hits.get(best)!) best = col;
  return best;
}

/**
 * Turn one master table into sessions. Nothing here knows the programme's
 * layout: the day/time columns and the level/section of every column are
 * discovered from the header rows (brief §3).
 */
export function extractSessions(table: RawTable, programme: string, tableIndex = 0): ParsedMaster {
  const warnings: Warning[] = [];
  const warn = (kind: Warning['kind'], message: string, cell?: RawCell) =>
    warnings.push({ kind, message, programme, row: cell?.row, col: cell?.col });

  const dayCol = bestColumn(table, (t) => parseDay(t) !== null);
  const timeCol = bestColumn(table, (t) => parseSlot(t) !== null);
  if (dayCol < 0 || timeCol < 0) {
    warn('structure', 'لم يتم العثور على عمود اليوم أو عمود الوقت في الجدول');
    return { programme, sessions: [], warnings };
  }
  const firstBodyRow = Math.min(
    ...table.cells.filter((c) => c.col === dayCol && parseDay(cellText(c)) !== null).map((c) => c.row),
  );
  // The day/time columns sit on the left in the Word masters and on the right in the civil workbook.
  const isData = (col: number) => col !== dayCol && col !== timeCol;
  const width = (start: number, end: number) =>
    table.colWidths.slice(start, end).reduce((a, b) => a + b, 0);
  const overlap = (a: { start: number; end: number }, b: { start: number; end: number }) =>
    width(Math.max(a.start, b.start), Math.min(a.end, b.end));

  // Header rows: level cells (`مستوى 2`) and section cells (bare numbers under `السكشن`).
  const header = table.cells.filter((c) => c.row < firstBodyRow && isData(c.col));
  const range = (c: RawCell) => ({ start: c.col, end: c.col + c.colSpan });
  const levels = header
    .filter((c) => levelNumber(cellText(c)) !== null)
    .map((c) => ({ ...range(c), level: levelNumber(cellText(c))! }));
  const levelAt = (r: { start: number; end: number }) => {
    let best: { level: number; w: number } | null = null;
    for (const l of levels) {
      const w = overlap(l, r);
      if (w > 0 && (!best || w > best.w)) best = { level: l.level, w };
    }
    return best?.level ?? null;
  };
  let columns: SectionCol[] = header
    .filter((c) => /^\d+$/.test(normalize(cellText(c))))
    .map((c) => ({
      ...range(c),
      level: levelAt(range(c)),
      section: Number(normalize(cellText(c))),
      width: width(c.col, c.col + c.colSpan),
    }));
  if (!columns.length) {
    // A programme with one group per level has no section row.
    columns = levels.map((l) => ({ ...l, section: 1, width: width(l.start, l.end) }));
  }
  if (!columns.length) warn('structure', 'لم يتم العثور على صف المستويات/السكاشن');
  const sectionsOfLevel = (level: number | null) => columns.filter((c) => c.level === level).length;

  // Day and slot of every body row. Rows without a label are sub-rows of the slot above.
  const origin: (RawCell | undefined)[][] = Array.from({ length: table.rowCount }, () => []);
  for (const c of table.cells) {
    for (let r = c.row; r < c.row + c.rowSpan; r++) {
      for (let k = c.col; k < c.col + c.colSpan; k++) origin[r][k] = c;
    }
  }
  const rowDay: (number | null)[] = [];
  const rowSlot: (number | null)[] = [];
  let day: number | null = null;
  let slot: number | null = null;
  // Rows under the timetable (signatures, notes) start with other text in the day column.
  const outside: boolean[] = [];
  let below = false;
  for (let r = firstBodyRow; r < table.rowCount; r++) {
    const dayCell = origin[r][dayCol];
    const d = dayCell ? parseDay(cellText(dayCell)) : null;
    if (d !== null) below = false;
    else if (dayCell && /\p{L}/u.test(cellText(dayCell))) below = true;
    outside[r] = below;
    if (below) {
      rowDay[r] = null;
      rowSlot[r] = null;
      continue;
    }
    if (d !== null && d !== day) {
      day = d;
      slot = null;
    }
    const timeCell = origin[r][timeCol];
    const label = timeCell ? tidy(cellText(timeCell)) : '';
    if (label) {
      const parsed = parseSlot(label);
      if (!parsed) warn('time-label', `تعذر قراءة الوقت «${label}»`, timeCell);
      else {
        if (d === null && slot !== null && parsed.slot < slot) {
          warn('day', `الصف ${r + 1}: الوقت «${label}» يسبق الصف الذي قبله بدون اسم يوم جديد`, timeCell);
        }
        slot = parsed.slot;
      }
    }
    rowDay[r] = day;
    rowSlot[r] = slot;
  }

  const sessions: Session[] = [];
  for (const cell of table.cells) {
    if (cell.row < firstBodyRow || !isData(cell.col) || outside[cell.row]) continue;
    const text = tidy(cellText(cell));
    if (parseDay(text) !== null) continue; // a second day column
    if (!/\p{L}/u.test(text)) continue; // empty or `-------`
    const cellDay = rowDay[cell.row];
    const slots = new Set<number>();
    for (let r = cell.row; r < cell.row + cell.rowSpan; r++) {
      if (rowDay[r] === cellDay && rowSlot[r] !== null) slots.add(rowSlot[r]!);
    }
    if (cellDay === null || !slots.size) {
      warn('day', `خلية بدون يوم/وقت: «${text}»`, cell);
      continue;
    }
    // A cell belongs to a section when it covers at least half of that section's column.
    let covered = columns.filter((c) => overlap(c, range(cell)) * 2 >= c.width);
    if (!covered.length) {
      const near = columns
        .map((c) => ({ c, w: overlap(c, range(cell)) }))
        .filter((x) => x.w > 0)
        .sort((a, b) => b.w - a.w)[0];
      if (near) covered = [near.c];
    }
    if (!covered.length) {
      warn('no-section', `خلية خارج أعمدة السكاشن: «${text}»`, cell);
      continue;
    }
    // A cell stretching over two levels is kept under the level it covers most.
    const level = covered[0].level;
    covered = covered.filter((c) => c.level === level);
    const parsed = parseCellText(cell.runs);
    if (parsed.type === 'ignore') continue;
    if (parsed.type === 'unknown') warn('unknown-type', `نوع الحصة غير معروف: «${text}»`, cell);
    sessions.push({
      id: '',
      programme,
      day: cellDay,
      slots: [...slots].sort((a, b) => a - b),
      level,
      sections: covered.map((c) => c.section).sort((a, b) => a - b),
      wholeLevel: covered.length === sectionsOfLevel(level),
      text,
      runs: cell.runs,
      parsed,
      scannedStaff: [],
      source: { table: tableIndex, row: cell.row, col: cell.col },
    });
  }

  return { programme, sessions: mergeAdjacent(sessions), warnings };
}

/** The same cell typed twice in consecutive slots (not merged in the master) is one session. */
function mergeAdjacent(sessions: Session[]): Session[] {
  const key = (s: Session) => [s.day, s.level, s.sections.join('+'), normalize(s.text)].join('|');
  const sorted = [...sessions].sort((a, b) => a.day - b.day || a.slots[0] - b.slots[0]);
  const last = new Map<string, Session>();
  const out: Session[] = [];
  for (const s of sorted) {
    const prev = last.get(key(s));
    if (prev && s.slots[0] - prev.slots[prev.slots.length - 1] <= 1) {
      prev.slots = [...new Set([...prev.slots, ...s.slots])].sort((a, b) => a - b);
      continue;
    }
    last.set(key(s), s);
    out.push(s);
  }
  out.forEach((s) => {
    s.id = `${s.programme}:${s.day}:${s.slots[0]}:L${s.level ?? '-'}:S${s.sections.join('+')}:${s.source.col}`;
  });
  return out;
}

/** All tables of one master file → one list of sessions. */
export function parseMaster(tables: RawTable[], programme: string): ParsedMaster {
  const out: ParsedMaster = { programme, sessions: [], warnings: [] };
  tables.forEach((t, i) => {
    const hasDays = t.cells.some((c) => parseDay(cellText(c)) !== null);
    if (!hasDays) return;
    const part = extractSessions(t, programme, i);
    out.sessions.push(...part.sessions);
    out.warnings.push(...part.warnings);
  });
  if (!out.sessions.length && !out.warnings.length) {
    out.warnings.push({ kind: 'structure', message: 'لم يتم العثور على جدول دراسي في الملف', programme });
  }
  return out;
}
