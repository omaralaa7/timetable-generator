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
  masters: StoredMaster[];
  roster: Staff[];
  decisions: Decisions;
  settings: Settings;
}

const KEY = 'individual-timetables-project-v1';

export function emptyProject(): Project {
  return {
    version: 1,
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
  return {
    version: 1,
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
