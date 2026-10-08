// Text preview of the generated sheets (before the Excel writer / UI exist).
//   node scripts/preview.ts              summary of everyone + issues
//   node scripts/preview.ts asmaa-radi   one person's courses and grid
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { generate, readMaster } from '../src/core/pipeline.ts';
import { SEED_ROSTER } from '../src/core/roster.ts';
import { DAYS, SLOT_LABELS } from '../src/core/slots.ts';

const dir = join(import.meta.dirname, '..', 'samples', 'masters');
const masters = [];
for (const name of readdirSync(dir).filter((f) => f.endsWith('.docx') || f.endsWith('.xlsx'))) {
  masters.push(await readMaster({ programme: name.split('_')[0], name, data: readFileSync(join(dir, name)) }));
}
const result = generate(masters, SEED_ROSTER);
const who = process.argv[2];

if (!who) {
  console.log('| name | teaching | office | supervision | advising | quality | counted / min | daily |');
  console.log('|---|---|---|---|---|---|---|---|');
  for (const { staff, totals: t, daily } of result.sheets) {
    console.log(`| ${staff.name} | ${t.teaching} | ${t.office} | ${t.supervision} | ${t.advising} | ${t.quality} | ${t.counted} / ${t.minimum} | ${daily.join(' ')} |`);
  }
  const kinds = new Map<string, number>();
  for (const i of result.issues) kinds.set(i.kind, (kinds.get(i.kind) ?? 0) + 1);
  console.log('\nissues:', Object.fromEntries(kinds));
  for (const i of result.issues) console.log(`- [${i.kind}] ${i.message}`);
  console.log('\ndetected, not in roster:', result.detected.map((d) => `${d.name} (${d.count})`).join('، '));
} else {
  const sheet = result.sheets.find((s) => s.staff.id === who);
  if (!sheet) throw new Error(`unknown staff id ${who}`);
  console.log(`${sheet.staff.name}\n`);
  for (const c of sheet.courses) console.log(`- ${c.name} | ${c.share} | ${c.weeks}`);
  console.log('');
  for (const it of sheet.items) {
    const time = `${SLOT_LABELS[it.slots[0]].slice(0, 5)}-${SLOT_LABELS[it.slots[it.slots.length - 1]].slice(6)}`;
    console.log(`| ${DAYS[it.day]} | ${time} | ${it.type} | ${it.text} |`);
  }
  console.log('\ntotals', sheet.totals, 'daily', sheet.daily.join(' '));
  for (const i of result.issues.filter((x) => x.staffIds.includes(who))) console.log(`- [${i.kind}] ${i.message}`);
}
