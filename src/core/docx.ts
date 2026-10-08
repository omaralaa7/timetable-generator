import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import type { RawCell, RawTable, Run } from './types.ts';

// Structural typing keeps this independent of xmldom's vs the browser's DOM types.
interface XNode {
  nodeName: string;
  childNodes: ArrayLike<XNode>;
  textContent: string | null;
  getAttribute?(name: string): string | null;
}

function children(node: XNode, name: string): XNode[] {
  const out: XNode[] = [];
  for (let i = 0; i < node.childNodes.length; i++) {
    if (node.childNodes[i].nodeName === name) out.push(node.childNodes[i]);
  }
  return out;
}

function child(node: XNode | undefined, name: string): XNode | undefined {
  return node ? children(node, name)[0] : undefined;
}

function descendants(node: XNode, name: string, stopAt?: string, out: XNode[] = []): XNode[] {
  for (let i = 0; i < node.childNodes.length; i++) {
    const c = node.childNodes[i];
    if (c.nodeName === name) out.push(c);
    if (c.nodeName !== stopAt) descendants(c, name, stopAt, out);
  }
  return out;
}

function val(node: XNode | undefined): string | null {
  return node?.getAttribute?.('w:val') ?? null;
}

function intVal(node: XNode | undefined, fallback: number): number {
  const n = Number(val(node));
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

function runText(r: XNode): string {
  let text = '';
  for (let i = 0; i < r.childNodes.length; i++) {
    const c = r.childNodes[i];
    if (c.nodeName === 'w:t') text += c.textContent ?? '';
    else if (c.nodeName === 'w:tab' || c.nodeName === 'w:br' || c.nodeName === 'w:cr') text += ' ';
    else if (c.nodeName === 'w:noBreakHyphen') text += '-';
  }
  return text;
}

function cellRuns(tc: XNode): Run[] {
  const runs: Run[] = [];
  descendants(tc, 'w:p').forEach((p, i) => {
    if (i > 0) runs.push({ text: '\n', underline: false });
    for (const r of descendants(p, 'w:r')) {
      const text = runText(r);
      if (!text) continue;
      const u = child(child(r, 'w:rPr'), 'w:u');
      const underline = !!u && val(u) !== 'none';
      const last = runs[runs.length - 1];
      if (last && last.underline === underline && last.text !== '\n') last.text += text;
      else runs.push({ text, underline });
    }
  });
  return runs;
}

function readTable(tbl: XNode): RawTable {
  const colWidths = children(child(tbl, 'w:tblGrid') ?? tbl, 'w:gridCol').map(
    (g) => Number(g.getAttribute?.('w:w')) || 1,
  );
  const cells: RawCell[] = [];
  // Origin cell still open for vertical merging, by its start column.
  let open = new Map<number, RawCell>();
  const rows = children(tbl, 'w:tr');
  rows.forEach((tr, row) => {
    const trPr = child(tr, 'w:trPr');
    let col = intVal(child(trPr, 'w:gridBefore'), 0);
    const next = new Map<number, RawCell>();
    for (const tc of children(tr, 'w:tc')) {
      const tcPr = child(tc, 'w:tcPr');
      const colSpan = intVal(child(tcPr, 'w:gridSpan'), 1);
      const vMerge = child(tcPr, 'w:vMerge');
      const continues = !!vMerge && val(vMerge) !== 'restart';
      const origin = continues ? open.get(col) : undefined;
      if (origin) {
        origin.rowSpan = row - origin.row + 1;
        next.set(col, origin);
      } else {
        const cell: RawCell = { row, col, rowSpan: 1, colSpan, runs: cellRuns(tc) };
        cells.push(cell);
        if (vMerge) next.set(col, cell);
      }
      col += colSpan;
    }
    open = next;
    while (colWidths.length < col) colWidths.push(1);
  });
  return { colWidths, rowCount: rows.length, cells };
}

/** Read every top-level table of `word/document.xml` as an occupancy grid. */
export function readDocumentXml(xml: string): RawTable[] {
  const doc = new DOMParser().parseFromString(xml, 'text/xml') as unknown as XNode;
  return descendants(doc, 'w:tbl', 'w:tbl').map(readTable);
}

export async function readDocx(data: ArrayBuffer | Uint8Array): Promise<RawTable[]> {
  const zip = await JSZip.loadAsync(data);
  const file = zip.file('word/document.xml');
  if (!file) throw new Error('الملف ليس مستند Word صالحاً (word/document.xml غير موجود)');
  return readDocumentXml(await file.async('string'));
}
