import ExcelJS from 'exceljs';
import type { RawCell, RawTable, Run } from './types.ts';

export interface XlsxMaster {
  /** Sheet names in workbook order. */
  sheets: string[];
  /** The sheet that was read. */
  sheet: string;
  table: RawTable;
}

function cellRuns(cell: ExcelJS.Cell): Run[] {
  const v = cell.value;
  if (v === null || v === undefined) return [];
  const whole = !!cell.font?.underline;
  if (typeof v === 'object' && 'richText' in v) {
    return v.richText
      .filter((r) => r.text)
      .map((r) => ({ text: r.text, underline: r.font ? !!r.font.underline : whole }));
  }
  const text =
    typeof v === 'object' && 'result' in v ? String(v.result ?? '') :
    typeof v === 'object' && 'text' in v ? String(v.text) :
    v instanceof Date ? v.toISOString() : String(v);
  return text ? [{ text, underline: whole }] : [];
}

function sheetText(ws: ExcelJS.Worksheet, rows = 8): string {
  let text = '';
  for (let r = 1; r <= Math.min(rows, ws.rowCount); r++) {
    ws.getRow(r).eachCell((c) => {
      text += ` ${cellRuns(c).map((x) => x.text).join('')}`;
    });
  }
  return `${ws.name} ${text}`;
}

/** The academic year mentioned in a sheet's name or heading (2026 for `2026-2027`), 0 if none. */
export function academicYear(text: string): number {
  const years = [...text.matchAll(/20\d\d/g)].map((m) => Number(m[0]));
  return years.length ? Math.min(...years) : 0;
}

function sheetYear(ws: ExcelJS.Worksheet): number {
  return academicYear(sheetText(ws));
}

/**
 * Read one sheet of an Excel master as an occupancy grid (brief §3.2).
 * Without `sheetName` the sheet of the newest academic year is used — the
 * civil workbook still carries last year's copy.
 */
export async function readXlsx(data: ArrayBuffer | Uint8Array, sheetName?: string): Promise<XlsxMaster> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(data as ArrayBuffer);
  const sheets = wb.worksheets.map((w) => w.name);
  const ws =
    wb.worksheets.find((w) => w.name === sheetName) ??
    [...wb.worksheets].sort((a, b) => sheetYear(b) - sheetYear(a))[0];
  if (!ws) throw new Error('ملف Excel لا يحتوي على أي ورقة');

  const cells: RawCell[] = [];
  const merges = new Map<string, { rowSpan: number; colSpan: number }>();
  for (const range of ws.model.merges ?? []) {
    const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(range);
    if (!m) continue;
    const c1 = ws.getColumn(m[1]).number;
    const c2 = ws.getColumn(m[3]).number;
    merges.set(`${m[1]}${m[2]}`, { rowSpan: Number(m[4]) - Number(m[2]) + 1, colSpan: c2 - c1 + 1 });
  }
  let colCount = 0;
  ws.eachRow((row, r) => {
    row.eachCell((cell, c) => {
      if (cell.isMerged && cell.master.address !== cell.address) return;
      const runs = cellRuns(cell);
      if (!runs.length) return;
      const span = merges.get(cell.address) ?? { rowSpan: 1, colSpan: 1 };
      // Hidden columns/rows hold leftovers the author no longer sees (the mechanics workbook
      // keeps another programme's cells in hidden columns) — read only what is visible.
      const visibleCol = Array.from({ length: span.colSpan }, (_, i) => c + i).some((x) => !ws.getColumn(x).hidden);
      const visibleRow = Array.from({ length: span.rowSpan }, (_, i) => r + i).some((x) => !ws.getRow(x).hidden);
      if (!visibleCol || !visibleRow) return;
      cells.push({ row: r - 1, col: c - 1, ...span, runs });
      colCount = Math.max(colCount, c - 1 + span.colSpan);
    });
  });
  const colWidths = Array.from({ length: colCount }, (_, i) => ws.getColumn(i + 1).width ?? 9);
  return { sheets, sheet: ws.name, table: { colWidths, rowCount: ws.rowCount, cells } };
}
