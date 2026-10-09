import { normalizeWithMap, tidy } from './normalize.ts';
import type { NameToken, ParsedText, Rank, Run, SessionType } from './types.ts';

const SEP = '\\s*(?:[./]+\\s*|\\s+|$)';
const TITLE_PREFIXES: [RegExp, NameToken['role'], Rank][] = [
  [new RegExp(`^\\.?\\s*ا\\s*\\.?\\s*م\\s*\\.?\\s*د${SEP}`), 'doctor', 'associate'],
  [new RegExp(`^\\.?\\s*ا\\s*\\.?\\s*د${SEP}`), 'doctor', 'professor'],
  // `.م.د` = ا.م.د with the alef lost
  [new RegExp(`^\\.?\\s*م\\s*\\.?\\s*د${SEP}`), 'doctor', 'associate'],
  [new RegExp(`^د${SEP}`), 'doctor', 'lecturer'],
  [new RegExp(`^م\\s*\\.?\\s*م${SEP}`), 'ta', 'ta'],
  [new RegExp(`^م${SEP}`), 'ta', 'ta'],
];

// The preparatory timetable abbreviates: `ت رياضيات 1` = تمرين, `م فيزياء 1` / `م برمجة` = معمل.
const KIND = /^(محاضره|مقرر اختياري|ت\s*\.?\s*عملي|عملي|معمل|تمرين|ورش|ت(?=\s)|م(?=\s))\s*/;
// The civil master writes the kind after the title: `رياضيات 3(تمرين)`.
const BRACKET_KIND = /^(تمرين|معمل|عملي)$/;
const CODE = /^([A-Z]{3})\s*([0-9A-Z]{1,4})$/;
const ROOM_WORD = /مدرج|قاعه|معمل|lab|no\s*\./i;
// A name stops where a room starts or another title follows: `(م. تامر م خطوط النقل`.
const NAME_STOP = /\s(?:(?:م|د)(?:\s*[./]\s*|\s+)\S|مدرج|قاعه|معمل|lab)/i;
const OPEN = /[(\[]/;
const CLOSE = /[)\]]/;

function detectType(n: string): SessionType {
  if (/انشطه (طلابيه|وندوات)/.test(n)) return 'ignore';
  if (/مشروع التخرج/.test(n)) return 'project';
  // Power writes its electives without the word `محاضرة`: `اختيارى 1 (ميكاترونكس) [ELP3E1]…`
  if (/محاضره|مقرر اختياري|^اختياري/.test(n)) return 'lecture';
  if (/عملي|معمل|تمرين|ورش|^[تم]\s/.test(n)) return 'section';
  return 'unknown';
}

interface Segment {
  start: number;
  end: number;
}

/** Split on a set of characters, keeping positions. */
function split(n: string, start: number, end: number, isSep: (i: number) => boolean): Segment[] {
  const out: Segment[] = [];
  let s = start;
  for (let i = start; i <= end; i++) {
    if (i === end || isSep(i)) {
      out.push({ start: s, end: i });
      s = i + 1;
    }
  }
  return out;
}

function trimRange(n: string, start: number, end: number, junk: RegExp): [number, number] {
  while (start < end && junk.test(n[start])) start++;
  while (end > start && junk.test(n[end - 1])) end--;
  return [start, end];
}

const titleAt = (text: string) => TITLE_PREFIXES.map(([re, role, rank]) => ({ m: re.exec(text), role, rank })).find((x) => x.m);

/** Read a session cell (brief §4): type, course, room and the staff names as written. */
export function parseCellText(runs: Run[]): ParsedText {
  const original = runs.map((r) => r.text).join('');
  const underlined: boolean[] = [];
  for (const r of runs) for (let i = 0; i < r.text.length; i++) underlined.push(r.underline);

  const { text: n, map } = normalizeWithMap(original);
  const slice = (start: number, end: number) =>
    start < end ? tidy(original.slice(map[start], map[end - 1] + 1)) : '';
  const underlineShare = (start: number, end: number) => {
    let letters = 0;
    let marked = 0;
    for (let i = start; i < end; i++) {
      if (!/\p{L}/u.test(n[i])) continue;
      letters++;
      if (underlined[map[i]]) marked++;
    }
    return letters ? marked / letters : 0;
  };

  // Brackets are too often unbalanced in the masters to be parsed as pairs.
  const segs = split(n, 0, n.length, (i) => OPEN.test(n[i]) || CLOSE.test(n[i]));
  const head = n.slice(0, segs[0].end);
  const kindMatch = KIND.exec(head);
  const names: NameToken[] = [];
  let code = '';
  let subtitle = '';
  let bracketKind = '';
  let lastNameEnd = -1;
  let lastHeadEnd = segs[0].end; // end of title / code / subtitle
  let titleSuffix = '';
  let tailStart = -1; // where the last name group was cut short by a room

  segs.forEach((seg, si) => {
    if (si === 0) return;
    const [s, e] = trimRange(n, seg.start, seg.end, /\s/);
    const body = n.slice(s, e);
    if (!body) return;
    const codeMatch = CODE.exec(body);
    if (codeMatch && codeMatch[1] !== 'LAB' && !code) {
      code = `${codeMatch[1]} ${codeMatch[2]}`;
      lastHeadEnd = seg.end;
      return;
    }
    if (/^\d+$/.test(body) && code && lastNameEnd < 0) {
      titleSuffix = ` ${body}`; // `رياضيات (MAT 101)3`
      return;
    }
    if (!kindMatch && !bracketKind && lastNameEnd < 0 && BRACKET_KIND.test(body)) {
      bracketKind = slice(s, e);
      lastHeadEnd = seg.end;
      return;
    }
    // Text after the closing bracket of the names is the room, even when it starts with `م` (= معمل).
    if (!OPEN.test(n[seg.start - 1] ?? '') && !CLOSE.test(n[seg.end] ?? '')) return;
    // Staff group: parts separated by `&` (or `-` before a title), the first named one starting with a title.
    const found: NameToken[] = [];
    let role: NameToken['role'] | null = null;
    const isNameSep = (i: number) => n[i] === '&' || (n[i] === '-' && !!titleAt(n.slice(i + 1, e).trimStart()));
    for (const part of split(n, s, e, isNameSep)) {
      const [a, b] = trimRange(n, part.start, part.end, /\s/);
      const partText = n.slice(a, b);
      const hit = titleAt(partText);
      if (!hit && !role) break;
      const prefixLen = hit?.m?.[0].length ?? 0;
      const stop = NAME_STOP.exec(n.slice(a + prefixLen, b));
      const nameEnd = stop ? a + prefixLen + stop.index : b;
      const [ns, ne] = trimRange(n, a + prefixLen, nameEnd, /[^\p{L}]/u);
      if (hit) role = hit.role;
      tailStart = stop ? nameEnd : -1;
      if (ns >= ne) continue; // a title with no name: `(م. ---)`
      found.push({
        name: slice(ns, ne),
        prefix: hit ? partText.slice(0, prefixLen).replace(/[\s./]+$/, '') : '',
        role: role!,
        rank: hit ? hit.rank : null,
        underline: underlineShare(ns, ne),
        staffId: null,
      });
    }
    if (role) {
      names.push(...found);
      lastNameEnd = seg.end;
    } else if (lastNameEnd < 0 && !ROOM_WORD.test(body) && /\p{L}/u.test(body)) {
      if (!subtitle) subtitle = slice(s, e);
      lastHeadEnd = seg.end;
    }
  });

  // Title: first segment without the type keyword.
  const kindLen = kindMatch ? kindMatch[0].length : 0;
  let titleEnd = segs[0].end;
  let roomStart = tailStart >= 0 ? tailStart : (lastNameEnd >= 0 ? lastNameEnd : lastHeadEnd) + 1;
  if (lastNameEnd < 0 && !code && !subtitle && !bracketKind) {
    // No brackets to lean on: the room starts at the first room word after the title.
    const m = ROOM_WORD.exec(head.slice(kindLen));
    if (m) {
      titleEnd = kindLen + m.index;
      roomStart = titleEnd;
    }
  }
  const [ts, te] = trimRange(n, kindLen, titleEnd, /[\s/\-.:]/);
  const [rs, re] = trimRange(n, Math.min(roomStart, n.length), n.length, /[\s/\-)\]&]/);
  let title = slice(ts, te);
  // A code typed without brackets stays in the first segment: `… المعلومات ELC 151 (د. …`, `رياضيات 3-MAT101-`
  const bare = /[\s\-]*\b(?!LAB)([A-Z]{3})\s*(\d[0-9A-Z]{2,3})\b[\s\-]*/.exec(title);
  if (bare) {
    if (!code) code = `${bare[1]} ${bare[2]}`;
    title = `${title.slice(0, bare.index)} ${title.slice(bare.index + bare[0].length)}`.trim();
  }
  if (bracketKind) title = title.replace(/^[مت]\s+/, ''); // civil prefixes its sections with `م` / `ت`
  title += titleSuffix;
  let room = slice(rs, re);
  const opens = (room.match(/\(/g) ?? []).length;
  const closes = (room.match(/\)/g) ?? []).length;
  if (opens > closes) room = room.endsWith('(') ? room.slice(0, -1).trim() : room + ')';

  return {
    type: detectType(n),
    kind: kindMatch ? slice(0, kindMatch[0].trimEnd().length) : bracketKind,
    title,
    subtitle,
    code,
    room,
    names,
  };
}
