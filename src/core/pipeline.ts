import JSZip from 'jszip';
import { readDoc } from './doc.ts';
import { readDocumentXml } from './docx.ts';
import { matchStaff, type Staff } from './roster.ts';
import { buildSheets, type Decisions, type Result } from './rules.ts';
import { parseMaster } from './sessions.ts';
import type { ParsedMaster } from './types.ts';
import { readXls } from './xls.ts';
import { readXlsx } from './xlsx.ts';

export interface MasterFile {
  /** Short programme label used in session ids, e.g. `comms`. */
  programme: string;
  name: string;
  data: ArrayBuffer | Uint8Array;
  /** Excel only: the sheet to read. Default: the newest academic year. */
  sheet?: string;
}

export interface ReadMaster extends ParsedMaster {
  /** Excel only: all sheet names and the one that was read. */
  sheets?: string[];
  sheet?: string;
}

type Format = 'docx' | 'xlsx' | 'doc' | 'xls' | 'pdf' | 'unknown';

/**
 * What the file really is, from its content — the extension is not trusted
 * (a `.doc` may be a renamed `.docx`, and the other way round).
 */
export async function detectFormat(data: ArrayBuffer | Uint8Array): Promise<Format> {
  const b = data instanceof Uint8Array ? data : new Uint8Array(data);
  const starts = (...sig: number[]) => sig.every((x, i) => b[i] === x);
  if (starts(0x25, 0x50, 0x44, 0x46)) return 'pdf';
  if (starts(0x50, 0x4b)) {
    const zip = await JSZip.loadAsync(b).catch(() => null);
    if (zip?.file('word/document.xml')) return 'docx';
    if (zip?.file('xl/workbook.xml')) return 'xlsx';
    return 'unknown';
  }
  if (starts(0xd0, 0xcf, 0x11, 0xe0)) {
    // Old Office container: Word has a `WordDocument` stream, Excel a `Workbook` one (names are UTF-16).
    const has = (name: string) => {
      const sig = [...name].flatMap((ch) => [ch.charCodeAt(0), 0]);
      for (let i = 0; i + sig.length <= b.length; i += 64) {
        if (sig.every((x, k) => b[i + k] === x)) return true;
      }
      return false;
    };
    if (has('WordDocument')) return 'doc';
    if (has('Workbook') || has('Book')) return 'xls';
  }
  return 'unknown';
}

/**
 * Read one uploaded master, whatever Word/Excel format it is in: .docx/.docm/.dotx, .xlsx/.xlsm/.xltx,
 * and the old binary .doc/.dot/.wps and .xls/.xlt/.et. PDF is refused with a clear message.
 */
export async function readMaster(file: MasterFile): Promise<ReadMaster> {
  const format = await detectFormat(file.data);
  if (format === 'docx') {
    const zip = await JSZip.loadAsync(file.data);
    return parseMaster(readDocumentXml(await zip.file('word/document.xml')!.async('string')), file.programme);
  }
  if (format === 'doc') return parseMaster(readDoc(file.data), file.programme);
  if (format === 'xlsx') {
    const { table, sheets, sheet } = await readXlsx(file.data, file.sheet);
    return { ...parseMaster([table], file.programme), sheets, sheet };
  }
  if (format === 'xls') {
    const { table, sheets, sheet } = readXls(file.data, file.sheet);
    return { ...parseMaster([table], file.programme), sheets, sheet };
  }
  if (format === 'pdf') throw new Error(`«${file.name}»: ملفات PDF غير مدعومة — برجاء رفع ملف Word أو Excel الأصلي`);
  throw new Error(`«${file.name}»: صيغة غير مدعومة — المطلوب ملف Word أو Excel (احفظه بصيغة ‎.docx أو ‎.xlsx)`);
}

export function generate(masters: ParsedMaster[], roster: Staff[], decisions: Decisions = {}): Result {
  const sessions = masters.flatMap((m) => m.sessions);
  for (const s of sessions) for (const t of s.parsed.names) t.staffId = null;
  matchStaff(sessions, roster);
  return buildSheets(sessions, roster, decisions);
}
