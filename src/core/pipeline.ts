import { readDocx } from './docx.ts';
import { matchStaff, type Staff } from './roster.ts';
import { buildSheets, type Decisions, type Result } from './rules.ts';
import { parseMaster } from './sessions.ts';
import type { ParsedMaster } from './types.ts';
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

/** Read one uploaded master. PDF and unknown formats are refused with a clear message. */
export async function readMaster(file: MasterFile): Promise<ReadMaster> {
  const ext = file.name.toLowerCase().split('.').pop();
  if (ext === 'docx') return parseMaster(await readDocx(file.data), file.programme);
  if (ext === 'xlsx') {
    const { table, sheets, sheet } = await readXlsx(file.data, file.sheet);
    return { ...parseMaster([table], file.programme), sheets, sheet };
  }
  if (ext === 'pdf') throw new Error(`«${file.name}»: ملفات PDF غير مدعومة — برجاء رفع ملف Word أو Excel الأصلي`);
  throw new Error(`«${file.name}»: صيغة غير مدعومة — المطلوب ‎.docx أو ‎.xlsx`);
}

export function generate(masters: ParsedMaster[], roster: Staff[], decisions: Decisions = {}): Result {
  const sessions = masters.flatMap((m) => m.sessions);
  for (const s of sessions) for (const t of s.parsed.names) t.staffId = null;
  matchStaff(sessions, roster);
  return buildSheets(sessions, roster, decisions);
}
