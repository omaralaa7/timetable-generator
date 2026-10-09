import * as CFB from 'cfb';
import type { RawCell, RawTable, Run } from './types.ts';
import { academicYear, type XlsxMaster } from './xlsx.ts';

// Reader for the old binary Excel format (.xls, and WPS Spreadsheets' .et): BIFF8 records
// inside an OLE container. It extracts what the timetable needs: cell text with underline,
// merged ranges, column widths and hidden rows/columns ([MS-XLS]).

interface Rec {
  id: number;
  data: Uint8Array;
}

function records(stream: Uint8Array): Rec[] {
  const out: Rec[] = [];
  for (let i = 0; i + 4 <= stream.length; ) {
    const id = stream[i] | (stream[i + 1] << 8);
    const len = stream[i + 2] | (stream[i + 3] << 8);
    out.push({ id, data: stream.subarray(i + 4, i + 4 + len) });
    i += 4 + len;
  }
  return out;
}

const u16 = (b: Uint8Array, i: number) => b[i] | (b[i + 1] << 8);
const u32 = (b: Uint8Array, i: number) => (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24)) >>> 0;

interface Text {
  text: string;
  /** Rich-text runs: from character `at`, font `font`. */
  fonts: { at: number; font: number }[];
}

/** The shared string table. A string may continue in the next record, restarting with its own flags byte. */
function sharedStrings(chunks: Uint8Array[]): Text[] {
  let c = 0;
  let i = 8; // cstTotal, cstUnique
  const total = u32(chunks[0], 4);
  const next = () => {
    c++;
    i = 0;
  };
  const byte = () => {
    if (i >= chunks[c].length) next();
    return chunks[c][i++];
  };
  const word = () => byte() | (byte() << 8);
  const out: Text[] = [];
  while (out.length < total && c < chunks.length) {
    if (i >= chunks[c].length) {
      if (c + 1 >= chunks.length) break;
      next();
    }
    const count = word();
    const flags = byte();
    let wide = (flags & 1) !== 0;
    const runs = flags & 8 ? word() : 0;
    const ext = flags & 4 ? word() | (word() << 16) : 0;
    let text = '';
    for (let k = 0; k < count; k++) {
      if (i >= chunks[c].length) {
        next();
        wide = (chunks[c][i++] & 1) !== 0; // continued character data repeats the flags byte
      }
      text += String.fromCharCode(wide ? chunks[c][i++] | (chunks[c][i++] << 8) : chunks[c][i++]);
    }
    const fonts: Text['fonts'] = [];
    for (let k = 0; k < runs; k++) fonts.push({ at: word(), font: word() });
    for (let k = 0; k < ext; k++) byte();
    out.push({ text, fonts });
  }
  return out;
}

function shortString(b: Uint8Array, at: number): string {
  const count = u16(b, at);
  const wide = (b[at + 2] & 1) !== 0;
  let text = '';
  for (let k = 0; k < count; k++) text += String.fromCharCode(wide ? u16(b, at + 3 + 2 * k) : b[at + 3 + k]);
  return text;
}

function rk(value: number): number {
  const n = value & 2 ? value >> 2 : new DataView(new Uint32Array([0, value & 0xfffffffc]).buffer).getFloat64(0, true);
  return value & 1 ? n / 100 : n;
}

interface Sheet {
  name: string;
  cells: Map<string, { row: number; col: number; runs: Run[] }>;
  merges: { r1: number; r2: number; c1: number; c2: number }[];
  widths: Map<number, number>;
  hiddenCols: Set<number>;
  hiddenRows: Set<number>;
  rowCount: number;
}

export function readXls(data: ArrayBuffer | Uint8Array, sheetName?: string): XlsxMaster {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const file = CFB.read(bytes, { type: 'array' });
  const entry = CFB.find(file, 'Workbook') ?? CFB.find(file, 'Book');
  if (!entry?.content) throw new Error('الملف ليس ملف Excel صالحاً');
  const stream = new Uint8Array(entry.content as ArrayLike<number>);
  const recs = records(stream);

  // ── Workbook globals: fonts, cell formats, sheet list, shared strings ──
  const underlineOfFont: boolean[] = [];
  const fontOfFormat: number[] = [];
  const bounds: { name: string; visible: boolean }[] = [];
  let strings: Text[] = [];
  let at = 0;
  for (; at < recs.length && recs[at].id !== 0x000a; at++) {
    const { id, data: d } = recs[at];
    if (id === 0x0031) {
      if (underlineOfFont.length === 4) underlineOfFont.push(false); // font index 4 does not exist
      underlineOfFont.push(d[10] !== 0);
    } else if (id === 0x00e0) fontOfFormat.push(u16(d, 0));
    else if (id === 0x0085) {
      const wide = (d[7] & 1) !== 0;
      let name = '';
      for (let k = 0; k < d[6]; k++) name += String.fromCharCode(wide ? u16(d, 8 + 2 * k) : d[8 + k]);
      if (d[5] === 0) bounds.push({ name, visible: (d[4] & 3) === 0 });
      else bounds.push({ name: '', visible: false }); // chart / macro sheet: keeps the order
    } else if (id === 0x00fc) {
      const chunks = [d];
      while (recs[at + 1]?.id === 0x003c) chunks.push(recs[++at].data);
      strings = sharedStrings(chunks);
    }
  }
  const underlined = (format: number, font?: number) => underlineOfFont[font ?? fontOfFormat[format] ?? 0] ?? false;
  const richRuns = (t: Text, format: number): Run[] => {
    const out: Run[] = [];
    const marks = [{ at: 0, font: fontOfFormat[format] ?? 0 }, ...t.fonts];
    marks.forEach((m, k) => {
      const piece = t.text.slice(m.at, marks[k + 1]?.at ?? t.text.length);
      if (piece) out.push({ text: piece, underline: underlined(format, m.font) });
    });
    return out;
  };

  // ── One substream per sheet, in the order of the sheet list ──
  const sheets: Sheet[] = [];
  let index = -1;
  let sheet: Sheet | null = null;
  let pendingFormula: { row: number; col: number; format: number } | null = null;
  for (at++; at < recs.length; at++) {
    const { id, data: d } = recs[at];
    if (id === 0x0809) {
      index++;
      sheet = { name: bounds[index]?.name ?? '', cells: new Map(), merges: [], widths: new Map(), hiddenCols: new Set(), hiddenRows: new Set(), rowCount: 0 };
      continue;
    }
    if (!sheet) continue;
    const put = (row: number, col: number, runs: Run[]) => {
      if (runs.some((r) => r.text)) sheet!.cells.set(`${row}:${col}`, { row, col, runs });
      sheet!.rowCount = Math.max(sheet!.rowCount, row + 1);
    };
    const plain = (row: number, col: number, format: number, text: string) => put(row, col, [{ text, underline: underlined(format) }]);
    if (id === 0x000a) {
      if (sheet.name && bounds[index]?.visible !== false) sheets.push(sheet);
      sheet = null;
    } else if (id === 0x00fd) {
      const t = strings[u32(d, 6)];
      if (t) put(u16(d, 0), u16(d, 2), richRuns(t, u16(d, 4)));
    } else if (id === 0x0204) plain(u16(d, 0), u16(d, 2), u16(d, 4), shortString(d, 6));
    else if (id === 0x027e) plain(u16(d, 0), u16(d, 2), u16(d, 4), String(rk(u32(d, 6))));
    else if (id === 0x00bd) {
      for (let k = 0; 4 + 6 * k + 6 <= d.length - 2; k++) plain(u16(d, 0), u16(d, 2) + k, u16(d, 4 + 6 * k), String(rk(u32(d, 6 + 6 * k))));
    } else if (id === 0x0203) plain(u16(d, 0), u16(d, 2), u16(d, 4), String(new DataView(d.buffer, d.byteOffset + 6, 8).getFloat64(0, true)));
    else if (id === 0x0006) {
      if (u16(d, 12) === 0xffff) pendingFormula = d[6] === 0 ? { row: u16(d, 0), col: u16(d, 2), format: u16(d, 4) } : null;
      else plain(u16(d, 0), u16(d, 2), u16(d, 4), String(new DataView(d.buffer, d.byteOffset + 6, 8).getFloat64(0, true)));
    } else if (id === 0x0207 && pendingFormula) {
      plain(pendingFormula.row, pendingFormula.col, pendingFormula.format, shortString(d, 0));
      pendingFormula = null;
    } else if (id === 0x00e5) {
      for (let k = 0; k < u16(d, 0); k++) {
        sheet.merges.push({ r1: u16(d, 2 + 8 * k), r2: u16(d, 4 + 8 * k), c1: u16(d, 6 + 8 * k), c2: u16(d, 8 + 8 * k) });
      }
    } else if (id === 0x007d) {
      for (let col = u16(d, 0); col <= Math.min(u16(d, 2), 255); col++) {
        sheet.widths.set(col, u16(d, 4) / 256);
        if (u16(d, 8) & 1) sheet.hiddenCols.add(col);
      }
    } else if (id === 0x0208 && u16(d, 12) & 0x20) sheet.hiddenRows.add(u16(d, 0));
  }
  if (!sheets.length) throw new Error('ملف Excel لا يحتوي على أي ورقة');

  // The sheet of the newest academic year, as for .xlsx.
  const head = (s: Sheet) => `${s.name} ${[...s.cells.values()].filter((c) => c.row < 8).map((c) => c.runs.map((r) => r.text).join('')).join(' ')}`;
  const chosen = sheets.find((s) => s.name === sheetName) ?? [...sheets].sort((a, b) => academicYear(head(b)) - academicYear(head(a)))[0];

  const cells: RawCell[] = [];
  let colCount = 0;
  const inside = (row: number, col: number) => chosen.merges.find((m) => row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2);
  for (const cell of chosen.cells.values()) {
    const m = inside(cell.row, cell.col);
    if (m && (m.r1 !== cell.row || m.c1 !== cell.col)) continue;
    const rowSpan = m ? m.r2 - m.r1 + 1 : 1;
    const colSpan = m ? m.c2 - m.c1 + 1 : 1;
    const visibleCol = Array.from({ length: colSpan }, (_, k) => cell.col + k).some((x) => !chosen.hiddenCols.has(x));
    const visibleRow = Array.from({ length: rowSpan }, (_, k) => cell.row + k).some((x) => !chosen.hiddenRows.has(x));
    if (!visibleCol || !visibleRow) continue;
    cells.push({ row: cell.row, col: cell.col, rowSpan, colSpan, runs: cell.runs });
    colCount = Math.max(colCount, cell.col + colSpan);
  }
  const table: RawTable = {
    colWidths: Array.from({ length: colCount }, (_, k) => chosen.widths.get(k) ?? 9),
    rowCount: Math.max(chosen.rowCount, ...cells.map((c) => c.row + c.rowSpan)),
    cells,
  };
  return { sheets: sheets.map((s) => s.name), sheet: chosen.name, table };
}
