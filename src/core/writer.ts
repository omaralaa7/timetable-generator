import ExcelJS from 'exceljs';
import { LOGO_FACULTY, LOGO_UNIVERSITY } from './logos.ts';
import type { GridItem, ItemType, StaffSheet } from './rules.ts';
import { DEFAULT_SETTINGS, type Settings } from './settings.ts';
import type { Rank } from './types.ts';

// Layout copied from samples/expected/asmaa_radi.xlsx (brief §7).
const COL_WIDTHS = [4.57, 20.57, 37.43, 42.71, 29.14, 39.86, 34.14, 35.14, 36.57, 41.14, 36.57, 30, 29.57, 2.57];
const DAY_NAMES = ['السبت', 'الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس'];
const TIME_LABELS = [
  '09:00 - 09:50', '09:50 - 10:40', '10:45 - 11:35', '11:35 - 12:25', '12:30 - 01:20',
  '01:20 - 02:10', '02:15 - 03:05', '03:05 - 03:55', '04:00 - 04:50', '04:50 - 05:40',
];
const RANK_LABEL: Record<Rank, string> = {
  professor: 'أستاذ دكتور', associate: 'أستاذ مساعد', lecturer: 'مدرس', ta: 'مدرس مساعد',
};
const FILL: Record<ItemType, string> = {
  lecture: 'FF92D050', supervision: 'FFFFFF00', section: 'FFFFFF00',
  office: 'FF99FFCC', advising: 'FF99FFCC', quality: 'FF99FFCC',
};
const GREY = 'FFF2F2F2';
const HEADER = 'FFE4DFEC';

type Line = 'thin' | 'medium' | 'thick';
type Edges = { top?: Line; bottom?: Line; left?: Line; right?: Line };

export function sheetFileName(sheet: StaffSheet): string {
  return `جدول فردي - ${sheet.staff.name.replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim()}.xlsx`;
}

/** Build one person's `.xlsx` in the layout of the hand-made sheets. */
export async function writeSheet(sheet: StaffSheet, settings: Settings = DEFAULT_SETTINGS): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('ورقة1', {
    views: [{ rightToLeft: true, zoomScale: 40 }],
    properties: { defaultRowHeight: 30 },
    pageSetup: {
      paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      margins: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    },
  });
  ws.columns = COL_WIDTHS.map((width) => ({ width }));

  const put = (
    range: string,
    value: ExcelJS.CellValue,
    opt: { size?: number; fill?: string; border?: Edges | Line; wrap?: boolean; align?: 'center' | 'right'; numFmt?: string } = {},
  ) => {
    const [from, to = from] = range.split(':');
    const a = ws.getCell(from);
    const b = ws.getCell(to);
    if (from !== to) ws.mergeCells(range);
    a.value = value;
    const edges: Edges = typeof opt.border === 'string'
      ? { top: opt.border, bottom: opt.border, left: opt.border, right: opt.border }
      : (opt.border ?? {});
    for (let r = Number(a.row); r <= Number(b.row); r++) {
      for (let c = Number(a.col); c <= Number(b.col); c++) {
        const cell = ws.getCell(r, c);
        cell.font = { name: 'Calibri', size: opt.size ?? 28, bold: true };
        cell.alignment = { horizontal: opt.align ?? 'center', vertical: 'middle', wrapText: opt.wrap ?? false, readingOrder: 'rtl' };
        if (opt.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opt.fill } };
        if (opt.numFmt) cell.numFmt = opt.numFmt;
        const border: Partial<ExcelJS.Borders> = {};
        if (edges.top && r === Number(a.row)) border.top = { style: edges.top };
        if (edges.bottom && r === Number(b.row)) border.bottom = { style: edges.bottom };
        // In a right-to-left sheet `left` is still the lower column letter.
        if (edges.left && c === Number(a.col)) border.left = { style: edges.left };
        if (edges.right && c === Number(b.col)) border.right = { style: edges.right };
        cell.border = border;
      }
    }
  };
  const col = (index: number) => String.fromCharCode(65 + index); // 0 → A
  const heights = (rows: Record<number, number>) => {
    for (const [r, h] of Object.entries(rows)) ws.getRow(Number(r)).height = h;
  };

  // ── Heading ──
  heights({ 3: 40, 4: 40, 5: 50, 6: 20, 7: 50, 8: 25, 9: 32, 10: 45 });
  put('D3:K4', 'جدول الأعباء التدريسية وملحقاتها', { size: 52 });
  put('D5:K5', settings.semester, { size: 36 });
  put('B7:G8', `الإسم: ${sheet.staff.name}`, { size: 60, fill: GREY });
  put('B10', 'القسم:', { size: 36, border: 'thin' });
  put('C10:E10', `${settings.department} (${sheet.staff.programme})`, { size: 36 });
  put('F10', 'الدرجة:', { size: 36, border: 'thin' });
  put('G10', RANK_LABEL[sheet.staff.rank], { size: 36 });

  // ── Courses ──
  put('J8:M9', 'أسماء المقررات المشارك فيها', { border: 'thick' });
  put('J10:K10', 'المقرر', { border: 'thick' });
  put('L10', 'نسبة المشاركة', { border: 'thick' });
  put('M10', 'الاسابيع', { border: 'thick' });
  const courses = sheet.courses.length ? sheet.courses : [{ key: '', name: '', share: null, weeks: '' }];
  courses.forEach((c, i) => {
    const r = 11 + i;
    ws.getRow(r).height = 48;
    put(`J${r}:K${r}`, c.name, { border: 'thick', wrap: true });
    put(`L${r}`, c.share, { border: 'thick', numFmt: '0%' });
    put(`M${r}`, c.weeks, { border: 'thick', numFmt: '@' });
  });

  // ── Weekly grid ──
  const head = 11 + courses.length + 1;
  heights({ [head - 1]: 30, [head]: 44 });
  put(`B${head}`, null, { size: 12, border: { top: 'thick', bottom: 'medium', left: 'thick', right: 'medium' } });
  TIME_LABELS.forEach((label, i) => put(`${col(2 + i)}${head}`, label, { border: { top: 'thick', left: 'medium', right: 'medium', bottom: 'medium' } }));
  put(`M${head}`, 'المجموع', { size: 32, border: 'medium' });
  DAY_NAMES.forEach((name, day) => {
    const r = head + 1 + day;
    ws.getRow(r).height = 115;
    put(`B${r}`, name, { size: 24, border: { top: 'medium', bottom: day === 5 ? 'thick' : 'medium', left: 'thick', right: 'medium' } });
    if (sheet.staff.secondmentDay === day) {
      put(`C${r}:L${r}`, 'انتــــــــــــــــــــــــــــــــــــــــــــــــــــداب', { size: 48, border: 'medium' });
      put(`M${r}`, null, { border: 'medium' });
      return;
    }
    const taken = new Set<number>();
    for (const item of sheet.items.filter((it) => it.day === day)) {
      for (const run of runs(item.slots.filter((s) => !taken.has(s)))) {
        run.forEach((s) => taken.add(s));
        const range = `${col(2 + run[0])}${r}:${col(2 + run[run.length - 1])}${r}`;
        const duty = item.type === 'office' || item.type === 'advising' || item.type === 'quality';
        put(range, item.text, { size: duty ? 28 : 24, fill: FILL[item.type], border: 'medium', wrap: true });
      }
    }
    for (let s = 0; s < 10; s++) if (!taken.has(s)) put(`${col(2 + s)}${r}`, null, { size: 24, border: 'medium' });
    put(`M${r}`, sheet.daily[day] || null, { border: 'medium' });
  });

  // ── Summary ──
  const sum = head + 8;
  heights({ [sum - 1]: 25, [sum]: 50, [sum + 1]: 31, [sum + 2]: 34, [sum + 5]: 40, [sum + 6]: 40 });
  put(`C${sum - 1}:M${sum - 1}`, null, { size: 11, border: 'medium' });
  const t = sheet.totals;
  const isTa = sheet.staff.rank === 'ta';
  const columns: [string, number][] = [
    [isTa ? 'تمرين/عملي' : 'المحاضرات', t.teaching],
    ['الساعات المكتبية', t.office],
    ['ساعات الإشراف', t.supervision],
    ['ساعات الإرشاد الأكاديمي', t.advising],
    ['الجــــــــــودة', t.quality],
  ];
  columns.forEach(([label, value], i) => {
    const a = col(2 + i * 2);
    const b = col(3 + i * 2);
    put(`${a}${sum}:${b}${sum}`, label, { size: 36, fill: HEADER, border: 'medium' });
    put(`${a}${sum + 1}:${b}${sum + 1}`, value, { border: 'medium' });
    // النصاب: the counted values; quality is left out (brief §5.4, §11.1).
    put(`${a}${sum + 2}:${b}${sum + 2}`, i < 4 ? value : null, { border: 'medium' });
  });
  put(`M${sum}`, 'المجموع', { border: 'medium' });
  put(`M${sum + 1}`, { formula: `SUM(C${sum + 1}:L${sum + 1})`, result: t.counted + t.quality }, { border: 'medium' });
  put(`B${sum + 2}`, 'النصاب', { size: 32, border: 'thick' });
  put(`M${sum + 2}`, { formula: `SUM(C${sum + 2}:J${sum + 2})`, result: t.counted }, { border: 'medium' });

  // ── Signatures ──
  ['B:D', 'E:G', 'I:J', 'L:M'].forEach((span, i) => {
    const [a, b] = span.split(':');
    put(`${a}${sum + 5}:${b}${sum + 5}`, settings.signatures[i]?.title ?? '');
    put(`${a}${sum + 6}:${b}${sum + 6}`, settings.signatures[i]?.name ?? '');
  });

  // ── Logos (right: university, left: faculty — the sheet is right-to-left) ──
  const university = wb.addImage({ base64: LOGO_UNIVERSITY, extension: 'png' });
  const faculty = wb.addImage({ base64: LOGO_FACULTY, extension: 'png' });
  ws.addImage(university, { tl: { col: 1.05, row: 2.1 }, ext: { width: 356, height: 126 } });
  ws.addImage(faculty, { tl: { col: 12.15, row: 2.05 }, ext: { width: 160, height: 143 } });

  ws.pageSetup.printArea = `A3:O${sum + 6}`;
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

/** Split slots into runs of consecutive slots; each run becomes one merged cell. */
function runs(slots: number[]): number[][] {
  const out: number[][] = [];
  for (const s of [...slots].sort((a, b) => a - b)) {
    const last = out[out.length - 1];
    if (last && s === last[last.length - 1] + 1) last.push(s);
    else out.push([s]);
  }
  return out;
}

export type { GridItem };
