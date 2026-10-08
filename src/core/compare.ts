import ExcelJS from 'exceljs';
import { normalize } from './normalize.ts';
import { parseDay, parseSlot } from './slots.ts';
import type { ItemType, StaffSheet } from './rules.ts';

/** A hand-made sheet from samples/expected, reduced to what can be compared (brief §10). */
export interface ExpectedSheet {
  courses: { name: string; share: number; weeks: string }[];
  /** `day:slot` → type and text. */
  cells: Map<string, { type: ItemType | 'secondment' | 'other'; text: string }>;
  totals: number[]; // teaching, office, supervision, advising, quality
}

const text = (v: ExcelJS.CellValue): string =>
  v === null || v === undefined ? '' :
  typeof v === 'object' && 'richText' in v ? v.richText.map((r) => r.text).join('') :
  typeof v === 'object' && 'result' in v ? String(v.result ?? '') : String(v);

const weeksKey = (s: string) => s.replace(/\s/g, '');
const lettersKey = (s: string) => normalize(s).replace(/[^\p{L}\p{N}]/gu, '');

function cellType(t: string): ItemType | 'secondment' | 'other' {
  const n = normalize(t);
  if (/^انت.*داب$/.test(n.replace(/\s/g, ''))) return 'secondment';
  if (/اشراف/.test(n)) return 'supervision';
  if (/مكتبيه/.test(n)) return 'office';
  if (/ارشاد/.test(n)) return 'advising';
  if (/جوده/.test(n)) return 'quality';
  if (/محاضره|مقرر اختياري|[A-Z]{3}\s*\d/.test(n)) return 'lecture';
  if (/عملي|تمرين|معمل/.test(n)) return 'section';
  return 'other';
}

export async function readExpected(data: ArrayBuffer | Uint8Array): Promise<ExpectedSheet> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(data as ArrayBuffer);
  const ws = wb.worksheets[0];
  const out: ExpectedSheet = { courses: [], cells: new Map(), totals: [] };
  let head = 0;
  for (let r = 11; r <= ws.rowCount; r++) {
    if (parseSlot(text(ws.getCell(r, 3).value))) {
      head = r;
      break;
    }
    const name = text(ws.getCell(r, 10).value).trim();
    if (name) out.courses.push({ name, share: Number(text(ws.getCell(r, 12).value)), weeks: weeksKey(text(ws.getCell(r, 13).value)) });
  }
  for (let r = head + 1; r <= ws.rowCount; r++) {
    const day = parseDay(text(ws.getCell(r, 2).value));
    if (day === null) {
      if (normalize(text(ws.getCell(r, 2).value)).includes('النصاب')) {
        out.totals = [3, 5, 7, 9, 11].map((c) => Number(text(ws.getCell(r - 1, c).value)) || 0);
        break;
      }
      continue;
    }
    for (let slot = 0; slot < 10; slot++) {
      const cell = ws.getCell(r, 3 + slot);
      const t = text((cell.isMerged ? cell.master : cell).value).trim();
      if (t) out.cells.set(`${day}:${slot}`, { type: cellType(t), text: t });
    }
  }
  return out;
}

export interface Diff {
  lines: string[];
  lectureSlots: { same: number; expected: number; generated: number };
  supervisionSlots: { same: number; expected: number; generated: number };
}

/** Readable differences between a generated sheet and the hand-made one. */
export function compare(sheet: StaffSheet, expected: ExpectedSheet): Diff {
  const lines: string[] = [];
  const mine = new Map<string, { type: ItemType; text: string }>();
  for (const it of sheet.items) for (const s of it.slots) mine.set(`${it.day}:${s}`, it);
  const teach = sheet.staff.rank === 'ta' ? 'section' : 'lecture';
  const count = (type: string) => {
    const exp = [...expected.cells].filter(([, c]) => c.type === type);
    const gen = [...mine].filter(([, c]) => c.type === type);
    return { same: exp.filter(([at]) => mine.get(at)?.type === type).length, expected: exp.length, generated: gen.length };
  };
  const lectureSlots = count(teach);
  const supervisionSlots = count('supervision');

  for (const c of expected.courses) {
    const code = /[A-Z]{3}\s*[0-9A-Z]{3}/.exec(c.name)?.[0].replace(/\s/g, '');
    const hit = sheet.courses.find((g) => (code && g.name.replace(/\s/g, '').includes(code)) || lettersKey(c.name).includes(lettersKey(g.name.replace(/\(.*$/, ''))));
    if (!hit) lines.push(`course missing: ${c.name} (${c.share}, ${c.weeks})`);
    else if (weeksKey(hit.weeks) !== c.weeks || hit.share !== c.share) lines.push(`course differs: ${c.name}: sheet ${c.share} / ${c.weeks}, generated ${hit.share} / ${weeksKey(hit.weeks)}`);
  }
  for (const [at, c] of expected.cells) {
    if (c.type !== 'lecture' && c.type !== 'supervision' && c.type !== 'section') continue;
    const g = mine.get(at);
    if (g?.type !== c.type) lines.push(`${at} sheet has ${c.type} «${c.text.slice(0, 60)}», generated ${g ? `${g.type} «${g.text.slice(0, 40)}»` : 'nothing'}`);
  }
  for (const [at, g] of mine) {
    if (g.type !== 'lecture' && g.type !== 'supervision' && g.type !== 'section') continue;
    const c = expected.cells.get(at);
    if (!c || (c.type !== 'lecture' && c.type !== 'supervision' && c.type !== 'section')) lines.push(`${at} generated ${g.type} «${g.text.slice(0, 60)}», sheet has ${c ? c.type : 'nothing'}`);
  }
  const t = sheet.totals;
  const got = [t.teaching, t.office, t.supervision, t.advising, t.quality];
  if (got.join() !== expected.totals.join()) lines.push(`totals: sheet ${expected.totals.join(' / ')}, generated ${got.join(' / ')}`);
  return { lines, lectureSlots, supervisionSlots };
}
