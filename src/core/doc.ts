import * as CFB from 'cfb';
import type { RawCell, RawTable, Run } from './types.ts';

// Reader for the old binary Word format (.doc, Word 97–2003). It extracts only what the
// timetable needs: tables (cell text, column widths, merged cells) and underline.
// Layout follows [MS-DOC]: text is in pieces (CLX), paragraph/row properties in PAPX FKPs,
// character properties in CHPX FKPs, a table row's shape in sprmTDefTable.

const SPRM_P_IN_TABLE = 0x2416;
const SPRM_P_TTP = 0x2417;
const SPRM_P_HUGE = 0x6646;
const SPRM_P_HUGE_OLD = 0x6645;
const SPRM_T_DEF_TABLE = 0xd608;
const SPRM_T_VERT_MERGE = 0xd62b;
const SPRM_C_UNDERLINE = 0x2a3e;

interface Sprm {
  op: number;
  data: Uint8Array;
}

function sprms(grpprl: Uint8Array): Sprm[] {
  const out: Sprm[] = [];
  let i = 0;
  while (i + 2 <= grpprl.length) {
    const op = grpprl[i] | (grpprl[i + 1] << 8);
    i += 2;
    let size: number;
    switch (op >> 13) {
      case 0: case 1: size = 1; break;
      case 2: case 4: case 5: size = 2; break;
      case 3: size = 4; break;
      case 7: size = 3; break;
      default:
        if (op === SPRM_T_DEF_TABLE || op === 0xd606) {
          size = (grpprl[i] | (grpprl[i + 1] << 8)) - 1;
          i += 2;
        } else {
          size = grpprl[i];
          i += 1;
        }
    }
    if (size < 0 || i + size > grpprl.length) break;
    out.push({ op, data: grpprl.subarray(i, i + size) });
    i += size;
  }
  return out;
}

interface Range {
  start: number; // file offset of the first character
  end: number;
  grpprl: Uint8Array;
}

/** Read a PlcBte (paragraph or character) and return every formatted run by file offset. */
function formattedRuns(doc: Uint8Array, table: Uint8Array, fc: number, lcb: number, paragraph: boolean): Range[] {
  const view = new DataView(table.buffer, table.byteOffset, table.byteLength);
  const dv = new DataView(doc.buffer, doc.byteOffset, doc.byteLength);
  const n = (lcb - 4) / 8;
  const out: Range[] = [];
  for (let k = 0; k < n; k++) {
    const page = view.getUint32(fc + 4 * (n + 1) + 4 * k, true) * 512;
    if (page + 512 > doc.length) continue;
    const crun = doc[page + 511];
    for (let i = 0; i < crun; i++) {
      const start = dv.getUint32(page + 4 * i, true);
      const end = dv.getUint32(page + 4 * (i + 1), true);
      let grpprl: Uint8Array = new Uint8Array(0);
      if (paragraph) {
        const at = doc[page + 4 * (crun + 1) + 13 * i] * 2;
        if (at) {
          let cb = doc[page + at];
          let from = page + at + 1;
          if (cb === 0) {
            cb = doc[page + at + 1];
            from += 1;
            cb *= 2;
          } else cb = cb * 2 - 1;
          grpprl = doc.subarray(from + 2, from + cb); // skip the style index
        }
      } else {
        const at = doc[page + 4 * (crun + 1) + i] * 2;
        if (at) grpprl = doc.subarray(page + at + 1, page + at + 1 + doc[page + at]);
      }
      out.push({ start, end, grpprl });
    }
  }
  return out.sort((a, b) => a.start - b.start);
}

function find(ranges: Range[], fc: number): Range | undefined {
  let lo = 0;
  let hi = ranges.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (fc < ranges[mid].start) hi = mid - 1;
    else if (fc >= ranges[mid].end) lo = mid + 1;
    else return ranges[mid];
  }
  return undefined;
}

interface Char {
  ch: string;
  fc: number;
}

function stream(file: CFB.CFB$Container, name: string): Uint8Array | null {
  const entry = CFB.find(file, name);
  return entry?.content ? new Uint8Array(entry.content as ArrayLike<number>) : null;
}

/** Read every table of a binary `.doc` as an occupancy grid. */
export function readDoc(data: ArrayBuffer | Uint8Array): RawTable[] {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const file = CFB.read(bytes, { type: 'array' });
  const doc = stream(file, 'WordDocument');
  if (!doc || doc.length < 0x1aa) throw new Error('الملف ليس مستند Word صالحاً');
  const dv = new DataView(doc.buffer, doc.byteOffset, doc.byteLength);
  const table = stream(file, dv.getUint16(0x0a, true) & 0x0200 ? '1Table' : '0Table');
  if (!table) throw new Error('الملف ليس مستند Word صالحاً');
  const extra = stream(file, 'Data');
  const tv = new DataView(table.buffer, table.byteOffset, table.byteLength);

  // ── Text: the piece table maps character positions to file offsets ──
  const chars: Char[] = [];
  let at = dv.getUint32(0x01a2, true);
  const clxEnd = at + dv.getUint32(0x01a6, true);
  while (at < clxEnd && table[at] === 1) at += 3 + tv.getUint16(at + 1, true);
  if (table[at] !== 2) throw new Error('تعذر قراءة نص المستند');
  const lcb = tv.getUint32(at + 1, true);
  const plc = at + 5;
  const pieces = (lcb - 4) / 12;
  const decoder = new TextDecoder('windows-1252');
  for (let i = 0; i < pieces; i++) {
    const count = tv.getUint32(plc + 4 * (i + 1), true) - tv.getUint32(plc + 4 * i, true);
    const raw = tv.getUint32(plc + 4 * (pieces + 1) + 8 * i + 2, true);
    const compressed = (raw & 0x40000000) !== 0;
    const fc = compressed ? (raw & 0x3fffffff) >> 1 : raw & 0x3fffffff;
    for (let k = 0; k < count; k++) {
      const pos = fc + (compressed ? k : 2 * k);
      if (pos + (compressed ? 1 : 2) > doc.length) break;
      const ch = compressed ? decoder.decode(doc.subarray(pos, pos + 1)) : String.fromCharCode(dv.getUint16(pos, true));
      chars.push({ ch, fc: pos });
    }
  }

  const paragraphs = formattedRuns(doc, table, dv.getUint32(0x0102, true), dv.getUint32(0x0106, true), true);
  const characters = formattedRuns(doc, table, dv.getUint32(0x00fa, true), dv.getUint32(0x00fe, true), false);
  const paragraphSprms = (fc: number): Sprm[] => {
    const list = sprms(find(paragraphs, fc)?.grpprl ?? new Uint8Array(0));
    // Large property sets (a wide table row) are stored in the Data stream.
    const huge = list.find((s) => s.op === SPRM_P_HUGE || s.op === SPRM_P_HUGE_OLD);
    if (!huge || !extra) return list;
    const offset = new DataView(huge.data.buffer, huge.data.byteOffset, 4).getUint32(0, true);
    if (offset + 2 > extra.length) return list;
    const size = extra[offset] | (extra[offset + 1] << 8);
    return sprms(extra.subarray(offset + 2, offset + 2 + size));
  };
  const underlined = (fc: number) => {
    const u = sprms(find(characters, fc)?.grpprl ?? new Uint8Array(0)).filter((s) => s.op === SPRM_C_UNDERLINE).pop();
    return !!u && u.data[0] !== 0;
  };

  // ── Rows: cells end with 0x07; the row ends with a 0x07 whose paragraph carries the row shape ──
  interface Row {
    cells: Run[][];
    edges: number[];
    merge: number[]; // per cell: 0 none, 1 continues the cell above, 2 starts a vertical merge
  }
  const tables: Row[][] = [];
  let rows: Row[] = [];
  let cells: Run[][] = [];
  let runs: Run[] = [];
  let hidden = 0; // inside a field's instruction
  const push = (text: string, underline: boolean) => {
    const last = runs[runs.length - 1];
    if (last && last.underline === underline && text !== '\n' && last.text !== '\n') last.text += text;
    else runs.push({ text, underline });
  };
  const closeTable = () => {
    if (rows.length) tables.push(rows);
    rows = [];
    cells = [];
  };
  for (const { ch, fc } of chars) {
    const code = ch.charCodeAt(0);
    if (code === 0x13) hidden++;
    else if (code === 0x14 || code === 0x15) hidden = Math.max(0, hidden - 1);
    else if (hidden) continue;
    else if (code === 0x07) {
      const props = paragraphSprms(fc);
      if (props.some((s) => s.op === SPRM_P_TTP && s.data[0])) {
        const def = props.filter((s) => s.op === SPRM_T_DEF_TABLE).pop();
        if (def && cells.length) {
          const d = new DataView(def.data.buffer, def.data.byteOffset, def.data.byteLength);
          const count = def.data[0];
          const edges = Array.from({ length: count + 1 }, (_, i) => d.getInt16(1 + 2 * i, true));
          const merge = Array.from({ length: count }, (_, i) => {
            const pos = 1 + 2 * (count + 1) + 20 * i;
            const flags = pos + 2 <= def.data.length ? d.getUint16(pos, true) : 0;
            return flags & 0x40 ? 2 : flags & 0x20 ? 1 : 0;
          });
          for (const s of props.filter((x) => x.op === SPRM_T_VERT_MERGE)) {
            if (s.data[1] < count) merge[s.data[1]] = s.data[2] === 3 ? 2 : s.data[2] === 1 ? 1 : 0;
          }
          rows.push({ cells: cells.slice(0, count), edges, merge });
        }
        cells = [];
      } else {
        cells.push(runs);
      }
      runs = [];
    } else if (code === 0x0d) {
      const inTable = paragraphSprms(fc).some((s) => s.op === SPRM_P_IN_TABLE && s.data[0]);
      if (inTable) push('\n', false);
      else {
        closeTable();
        runs = [];
      }
    } else if (code === 0x0b || code === 0x09 || code === 0x0c) push(' ', false);
    else if (code === 0x1e) push('-', underlined(fc));
    else if (code >= 0x20) push(ch, underlined(fc));
  }
  closeTable();

  // ── Grid: rows have their own cell edges; the union of all edges is the column grid ──
  return tables.map((tableRows): RawTable => {
    const all = [...new Set(tableRows.flatMap((r) => r.edges))].sort((a, b) => a - b);
    const grid = all.filter((x, i) => i === 0 || x - all[i - 1] > 20); // edges a hair apart are the same edge
    const column = (x: number) => grid.reduce((best, g, i) => (Math.abs(g - x) < Math.abs(grid[best] - x) ? i : best), 0);
    const out: RawCell[] = [];
    const open = new Map<number, RawCell>();
    tableRows.forEach((row, r) => {
      const next = new Map<number, RawCell>();
      row.cells.forEach((cellRuns, i) => {
        const col = column(row.edges[i]);
        const colSpan = Math.max(1, column(row.edges[i + 1]) - col);
        const above = row.merge[i] === 1 ? open.get(col) : undefined;
        if (above) {
          above.rowSpan = r - above.row + 1;
          next.set(col, above);
          return;
        }
        while (cellRuns.length && cellRuns[cellRuns.length - 1].text === '\n') cellRuns.pop();
        const cell: RawCell = { row: r, col, rowSpan: 1, colSpan, runs: cellRuns };
        out.push(cell);
        if (row.merge[i]) next.set(col, cell);
      });
      open.clear();
      next.forEach((v, k) => open.set(k, v));
    });
    const colWidths = grid.slice(1).map((x, i) => Math.max(1, x - grid[i]));
    return { colWidths, rowCount: tableRows.length, cells: out };
  });
}
