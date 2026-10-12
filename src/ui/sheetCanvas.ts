import { LOGO_FACULTY, LOGO_UNIVERSITY } from '../core/logos.ts';
import type { PdfPage } from '../core/pdf.ts';
import type { ItemType, StaffSheet } from '../core/rules.ts';
import type { Settings } from '../core/settings.ts';
import type { Rank } from '../core/types.ts';

// Draws one person's sheet on a canvas in the layout of the Excel sheet (A4 landscape),
// for the PDF download. The browser lays out the Arabic text, so it is shaped correctly.

const W = 2480; // ≈ 212 dpi on A4 landscape
const H = 1754;
const M = 70;
const DAY_NAMES = ['السبت', 'الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس'];
const TIME_LABELS = [
  '09:00 - 09:50', '09:50 - 10:40', '10:45 - 11:35', '11:35 - 12:25', '12:30 - 01:20',
  '01:20 - 02:10', '02:15 - 03:05', '03:05 - 03:55', '04:00 - 04:50', '04:50 - 05:40',
];
const RANK_LABEL: Record<Rank, string> = { professor: 'أستاذ دكتور', associate: 'أستاذ مساعد', lecturer: 'مدرس', ta: 'مدرس مساعد' };
const FILL: Record<ItemType, string> = {
  lecture: '#92D050', supervision: '#FFFF00', section: '#FFFF00', office: '#99FFCC', advising: '#99FFCC', quality: '#99FFCC',
  activity: '#D9D9D9',
};
const FONT = '"Segoe UI", Tahoma, "Noto Sans Arabic", "Geeza Pro", Arial, sans-serif';

function image(base64: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('logo'));
    img.src = `data:image/png;base64,${base64}`;
  });
}

interface TextOptions {
  size: number;
  bold?: boolean;
  ltr?: boolean;
  color?: string;
  pad?: number;
}

export async function drawSheet(sheet: StaffSheet, settings: Settings): Promise<HTMLCanvasElement> {
  const canvas = document.createElement('canvas');
  canvas.width = W;
  canvas.height = H;
  const ctx = canvas.getContext('2d')!;
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, W, H);
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';

  const box = (x: number, y: number, w: number, h: number, fill?: string, line = 3) => {
    if (fill) {
      ctx.fillStyle = fill;
      ctx.fillRect(x, y, w, h);
    }
    if (line) {
      ctx.lineWidth = line;
      ctx.strokeStyle = '#000';
      ctx.strokeRect(x, y, w, h);
    }
  };
  /** Centre text in a box: wrap on spaces and shrink the font until it fits. */
  const text = (value: string | number | null, x: number, y: number, w: number, h: number, o: TextOptions) => {
    const str = value === null ? '' : String(value);
    if (!str) return;
    const pad = o.pad ?? 12;
    ctx.direction = o.ltr ? 'ltr' : 'rtl';
    ctx.fillStyle = o.color ?? '#000';
    for (let size = o.size; size >= 12; size -= 2) {
      ctx.font = `${o.bold === false ? 400 : 700} ${size}px ${FONT}`;
      const lines: string[] = [];
      let line = '';
      for (const word of str.split(/\s+/)) {
        const next = line ? `${line} ${word}` : word;
        if (line && ctx.measureText(next).width > w - 2 * pad) {
          lines.push(line);
          line = word;
        } else line = next;
      }
      lines.push(line);
      const lineHeight = size * 1.32;
      const fits = lines.length * lineHeight <= h - 6 && lines.every((l) => ctx.measureText(l).width <= w - 2 * pad);
      if (!fits && size > 12) continue;
      lines.forEach((l, i) => ctx.fillText(l, x + w / 2, y + h / 2 + (i - (lines.length - 1) / 2) * lineHeight, w - pad));
      return;
    }
  };
  const cell = (value: string | number | null, x: number, y: number, w: number, h: number, o: TextOptions & { fill?: string; line?: number }) => {
    box(x, y, w, h, o.fill, o.line);
    text(value, x, y, w, h, o);
  };

  // ── Heading ──
  const [university, faculty] = await Promise.all([image(LOGO_UNIVERSITY), image(LOGO_FACULTY)]);
  ctx.drawImage(university, W - M - 440, M - 10, 440, 440 * (university.height / university.width));
  ctx.drawImage(faculty, M, M - 20, 190, 190 * (faculty.height / faculty.width));
  text('جدول الأعباء التدريسية وملحقاتها', 600, M - 10, W - 1200, 100, { size: 66 });
  text(settings.semester, 600, M + 90, W - 1200, 70, { size: 42 });

  // ── Courses (left) and who the sheet is for (right) ──
  const top = M + 200;
  const courses = sheet.courses.length ? sheet.courses : [{ key: '', name: '', share: null as number | null, weeks: '' }];
  const coursesW = 900;
  const rowH = Math.max(34, Math.min(50, 400 / courses.length));
  const [cName, cShare, cWeeks] = [540, 200, 160];
  cell('أسماء المقررات المشارك فيها', M, top, coursesW, 54, { size: 30 });
  cell('الاسابيع', M, top + 54, cWeeks, 48, { size: 26 });
  cell('نسبة المشاركة', M + cWeeks, top + 54, cShare, 48, { size: 26 });
  cell('المقرر', M + cWeeks + cShare, top + 54, cName, 48, { size: 26 });
  courses.forEach((c, i) => {
    const y = top + 102 + i * rowH;
    cell(c.weeks, M, y, cWeeks, rowH, { size: 26, ltr: true });
    cell(c.share === null ? '' : `${Math.round(c.share * 100)}%`, M + cWeeks, y, cShare, rowH, { size: 26, ltr: true });
    cell(c.name, M + cWeeks + cShare, y, cName, rowH, { size: 26 });
  });
  const blockH = 102 + courses.length * rowH;
  const infoX = M + coursesW + 40;
  const infoW = W - M - infoX;
  cell(`الإسم: ${sheet.staff.name}`, infoX, top, infoW, 120, { size: 60, fill: '#F2F2F2', line: 0 });
  // Right to left: القسم | programme | الدرجة | rank
  const lineY = top + 150;
  cell('القسم:', W - M - 170, lineY, 170, 70, { size: 38, line: 2 });
  text(`${settings.department} (${sheet.staff.programme})`, infoX + 420, lineY, infoW - 590 - 20, 70, { size: 38 });
  cell('الدرجة:', infoX + 240, lineY, 170, 70, { size: 38, line: 2 });
  text(RANK_LABEL[sheet.staff.rank], infoX, lineY, 240, 70, { size: 38 });

  // ── Weekly grid (the day column is on the right) ──
  const gridTop = top + Math.max(blockH, 230) + 28;
  const bottomH = 300; // summary + signatures
  const gridH = H - M - bottomH - gridTop;
  const headH = 54;
  const dayH = (gridH - headH) / 6;
  const dayW = 190;
  const totalW = 150;
  const slotW = (W - 2 * M - dayW - totalW) / 10;
  const slotX = (slot: number) => W - M - dayW - (slot + 1) * slotW; // slot 0 is next to the day column
  cell('', W - M - dayW, gridTop, dayW, headH, { size: 20 });
  TIME_LABELS.forEach((label, i) => cell(label, slotX(i), gridTop, slotW, headH, { size: 26, ltr: true }));
  cell('المجموع', M, gridTop, totalW, headH, { size: 28 });
  DAY_NAMES.forEach((name, day) => {
    const y = gridTop + headH + day * dayH;
    cell(name, W - M - dayW, y, dayW, dayH, { size: 30 });
    if (sheet.staff.secondmentDay === day) {
      cell('انتـــــــــــــــــــــــــــــــــــــــــــــداب', slotX(9), y, slotW * 10, dayH, { size: 54 });
      cell('', M, y, totalW, dayH, { size: 30 });
      return;
    }
    const taken = new Set<number>();
    for (const item of sheet.items.filter((it) => it.day === day)) {
      const free = item.slots.filter((s) => !taken.has(s)).sort((a, b) => a - b);
      for (let i = 0; i < free.length; ) {
        let j = i;
        while (j + 1 < free.length && free[j + 1] === free[j] + 1) j++;
        const span = j - i + 1;
        cell(item.text, slotX(free[j]), y, slotW * span, dayH, { size: 27, fill: FILL[item.type] });
        for (let k = i; k <= j; k++) taken.add(free[k]);
        i = j + 1;
      }
    }
    for (let s = 0; s < 10; s++) if (!taken.has(s)) box(slotX(s), y, slotW, dayH);
    cell(sheet.daily[day] || '', M, y, totalW, dayH, { size: 32 });
  });

  // ── Summary ──
  const sumTop = gridTop + gridH + 16;
  const t = sheet.totals;
  const pairs: [string, number][] = [
    [sheet.staff.rank === 'ta' ? 'تمرين/عملي' : 'المحاضرات', t.teaching],
    ['الساعات المكتبية', t.office],
    ['ساعات الإشراف', t.supervision],
    ['ساعات الإرشاد الأكاديمي', t.advising],
    ['الجودة', t.quality],
  ];
  pairs.forEach(([label, value], i) => {
    const x = slotX(2 * i + 1);
    cell(label, x, sumTop, slotW * 2, 52, { size: 32, fill: '#E4DFEC' });
    cell(value, x, sumTop + 52, slotW * 2, 44, { size: 30 });
    // النصاب: the counted values; quality is left out.
    cell(i < 4 ? value : '', x, sumTop + 96, slotW * 2, 44, { size: 30 });
  });
  cell('المجموع', M, sumTop, totalW, 52, { size: 28 });
  cell(t.counted + t.quality, M, sumTop + 52, totalW, 44, { size: 30 });
  cell(t.counted, M, sumTop + 96, totalW, 44, { size: 30 });
  cell('النصاب', W - M - dayW, sumTop + 96, dayW, 44, { size: 30 });

  // ── Signatures (first one on the right) ──
  const signW = (W - 2 * M) / 4;
  settings.signatures.slice(0, 4).forEach((s, i) => {
    const x = W - M - (i + 1) * signW;
    text(s.title, x, sumTop + 160, signW, 46, { size: 30 });
    text(s.name, x, sumTop + 206, signW, 46, { size: 28 });
  });
  return canvas;
}

/** The drawn sheet as one PDF page. */
export async function sheetPage(sheet: StaffSheet, settings: Settings): Promise<PdfPage> {
  const canvas = await drawSheet(sheet, settings);
  const blob = await new Promise<Blob>((resolve, reject) =>
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('تعذر إنشاء صورة الجدول'))), 'image/jpeg', 0.92));
  return { jpeg: new Uint8Array(await blob.arrayBuffer()), width: canvas.width, height: canvas.height };
}
