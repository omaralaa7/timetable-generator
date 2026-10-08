// Debug table of the sessions found in the Word masters (brief §12 step 1).
//   node scripts/dump-sessions.ts                 every session
//   node scripts/dump-sessions.ts --names         distinct staff names as written
//   node scripts/dump-sessions.ts --staff asmaa-radi
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { readMaster } from '../src/core/pipeline.ts';
import { matchStaff, SEED_ROSTER } from '../src/core/roster.ts';
import { DAYS, SLOT_LABELS } from '../src/core/slots.ts';
import { nameKey } from '../src/core/normalize.ts';
import type { Session } from '../src/core/types.ts';

const dir = join(import.meta.dirname, '..', 'samples', 'masters');
const args = process.argv.slice(2);
const staffId = args.includes('--staff') ? args[args.indexOf('--staff') + 1] : null;

const sessions: Session[] = [];
for (const file of readdirSync(dir).filter((f) => f.endsWith('.docx') || f.endsWith('.xlsx'))) {
  const master = await readMaster({ programme: file.split('_')[0], name: file, data: readFileSync(join(dir, file)) });
  matchStaff(master.sessions, SEED_ROSTER);
  sessions.push(...master.sessions);
  console.log(`\n${file}: ${master.sessions.length} sessions, ${master.warnings.length} warnings`);
  for (const w of master.warnings) console.log(`  ! [${w.kind}] row ${w.row} col ${w.col}: ${w.message}`);
}

if (args.includes('--names')) {
  const seen = new Map<string, { written: Set<string>; n: number; staff: string | null; role: string }>();
  for (const s of sessions) {
    for (const t of s.parsed.names) {
      const k = `${t.role}:${nameKey(t.name)}`;
      const e = seen.get(k) ?? { written: new Set(), n: 0, staff: t.staffId, role: t.role };
      e.written.add(`${t.prefix} ${t.name}`.trim());
      e.n++;
      seen.set(k, e);
    }
  }
  for (const e of [...seen.values()].sort((a, b) => b.n - a.n)) {
    console.log(`${String(e.n).padStart(3)}  ${(e.staff ?? '—').padEnd(16)} ${[...e.written].join(' | ')}`);
  }
} else {
  const rows = sessions.filter(
    (s) => !staffId || s.parsed.names.some((t) => t.staffId === staffId) || s.scannedStaff.includes(staffId),
  );
  rows.sort((a, b) => a.day - b.day || a.slots[0] - b.slots[0] || a.programme.localeCompare(b.programme));
  console.log('\n| day | slots | time | prog | L | sec | type | kind | title | code | names | room |');
  console.log('|---|---|---|---|---|---|---|---|---|---|---|---|');
  for (const s of rows) {
    const p = s.parsed;
    const time = `${SLOT_LABELS[s.slots[0]].slice(0, 5)}-${SLOT_LABELS[s.slots[s.slots.length - 1]].slice(6)}`;
    const names = p.names
      .map((t) => `${t.prefix} ${t.name}${t.underline > 0.5 ? ' [U]' : ''}${t.staffId ? ` =${t.staffId}` : ''}`)
      .join(' & ');
    console.log(
      `| ${DAYS[s.day]} | ${s.slots.map((x) => x + 1).join(',')} | ${time} | ${s.programme} | ${s.level ?? '?'} | ${
        s.wholeLevel ? 'all' : s.sections.join('+')
      } | ${p.type} | ${p.kind} | ${p.title}${p.subtitle ? ` «${p.subtitle}»` : ''} | ${p.code} | ${names}${
        s.scannedStaff.length ? ` (scan: ${s.scannedStaff.join(',')})` : ''
      } | ${p.room} |`,
    );
  }
}
