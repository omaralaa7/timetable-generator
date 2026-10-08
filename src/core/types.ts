/** A piece of cell text with its underline flag (needed for the weeks rule, brief §5.2). */
export interface Run {
  text: string;
  underline: boolean;
}

/** One origin cell of a master table, already resolved to grid coordinates. */
export interface RawCell {
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
  runs: Run[];
}

/** A master table as an occupancy grid, independent of the source format (docx / xlsx). */
export interface RawTable {
  colWidths: number[];
  rowCount: number;
  cells: RawCell[];
}

export type Rank = 'professor' | 'associate' | 'lecturer' | 'ta';

export type SessionType = 'lecture' | 'section' | 'project' | 'ignore' | 'unknown';

/** A staff name as written in a cell. */
export interface NameToken {
  /** The name as written, without the title prefix. */
  name: string;
  /** Title prefix as written (normalised), e.g. `ا.م.د`. Empty when the name had none. */
  prefix: string;
  role: 'doctor' | 'ta';
  rank: Rank | null;
  /** Share of the name's letters that sit in underlined runs (0…1). */
  underline: number;
  /** Roster id, filled by matchStaff(). */
  staffId: string | null;
}

export interface ParsedText {
  type: SessionType;
  /** Type keyword as written: `محاضرة`, `ت عملى`, `تمرين`, `مقرر اختيارى`… */
  kind: string;
  /** Course title as written (tatweel removed). */
  title: string;
  /** Extra title in its own brackets, e.g. the real name of an elective. */
  subtitle: string;
  /** Course code in canonical form (`ELC 151`), or '' when the cell has none. */
  code: string;
  room: string;
  names: NameToken[];
}

export interface Session {
  id: string;
  programme: string;
  /** 0 = Saturday … 5 = Thursday. */
  day: number;
  /** Slot indexes 0…9, sorted. */
  slots: number[];
  level: number | null;
  sections: number[];
  /** True when the cell covers every section of its level. */
  wholeLevel: boolean;
  text: string;
  runs: Run[];
  parsed: ParsedText;
  /** Roster ids found by scanning the text for aliases but missed by the name tokens. */
  scannedStaff: string[];
  source: { table: number; row: number; col: number };
}

export interface Warning {
  kind: 'time-label' | 'day' | 'no-section' | 'unknown-type' | 'structure';
  message: string;
  programme: string;
  row?: number;
  col?: number;
}

export interface ParsedMaster {
  programme: string;
  sessions: Session[];
  warnings: Warning[];
}
