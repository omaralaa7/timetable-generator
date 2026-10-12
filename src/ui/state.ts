import { SEED_ROSTER, type Staff } from '../core/roster.ts';
import type { Decisions } from '../core/rules.ts';
import { DEFAULT_SETTINGS, type Settings } from '../core/settings.ts';
import type { Session, Warning } from '../core/types.ts';

export interface StoredMaster {
  /** Programme label = file name without extension; part of every session id. */
  programme: string;
  fileName: string;
  sheet?: string;
  sheets?: string[];
  sessions: Session[];
  warnings: Warning[];
}

/** Everything the user has done. Saved in the browser and exportable as a project file (brief §8). */
export interface Project {
  version: 1;
  dataVersion?: number;
  masters: StoredMaster[];
  roster: Staff[];
  decisions: Decisions;
  settings: Settings;
}

const KEY = 'individual-timetables-project-v1';

/**
 * Raised whenever the built-in staff data or the way masters are read changes. A project saved by an
 * older version then gets the new staff list and asks for the masters again, so nobody keeps working
 * on stale data. Review choices and any staff the user added are kept; hand edits of a person's grid are reset.
 */
export const DATA_VERSION = 7;
/** True when the saved project came from an older version and was refreshed on load. */
export let refreshed = false;

export function emptyProject(): Project {
  return {
    version: 1,
    dataVersion: DATA_VERSION,
    masters: [],
    roster: structuredClone(SEED_ROSTER),
    decisions: {},
    settings: structuredClone(DEFAULT_SETTINGS),
  };
}

/** Accept a parsed project file, filling anything missing from the defaults. */
export function normaliseProject(raw: unknown): Project {
  const p = raw as Partial<Project> | null;
  if (!p || typeof p !== 'object' || !Array.isArray(p.roster)) throw new Error('ملف المشروع غير صالح');
  const base = emptyProject();
  if (p.dataVersion !== DATA_VERSION) {
    refreshed = true;
    const builtIn = new Set(base.roster.map((x) => x.id));
    const added = p.roster.filter((x) => typeof x?.id === 'string' && x.id.startsWith('staff-') && !builtIn.has(x.id));
    // Review choices still apply. Hand-edited cells, frozen duty blocks and edited course tables were made
    // on top of the old rules (they can hold advising or office hours the new rules no longer give), so they go.
    const { starts, supervisor, names, dismissed } = p.decisions ?? {};
    return { ...base, roster: [...base.roster, ...added], decisions: { starts, supervisor, names, dismissed }, settings: { ...base.settings, ...p.settings } };
  }
  return {
    version: 1,
    dataVersion: DATA_VERSION,
    masters: Array.isArray(p.masters) ? p.masters : [],
    roster: p.roster,
    decisions: p.decisions ?? {},
    settings: { ...base.settings, ...p.settings },
  };
}

export function loadProject(): Project {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw) return normaliseProject(JSON.parse(raw));
  } catch {
    // A corrupt save must not lock the user out; start clean.
  }
  return emptyProject();
}

export function saveProject(project: Project): boolean {
  try {
    localStorage.setItem(KEY, JSON.stringify(project));
    return true;
  } catch {
    return false;
  }
}
