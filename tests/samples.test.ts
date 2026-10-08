// Generated sheets vs the ten hand-made ones (brief §10). The samples come from an
// older master and from programmes that are still PDF-only, so this prints a
// readable diff per person and only asserts what must hold.
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { test } from 'node:test';
import { compare, readExpected } from '../src/core/compare.ts';
import { generate, readMaster } from '../src/core/pipeline.ts';
import { SEED_ROSTER } from '../src/core/roster.ts';

const samples = join(import.meta.dirname, '..', 'samples');
const FILES: Record<string, string> = {
  'asmaa-radi': 'asmaa_radi', 'ahmed-salem': 'ahmed_salem', 'mai-helmy': 'mai_helmy', 'hend-elsayed': 'hend_elsayed',
  'ahmed-elsayed': 'ahmed_elsayed', 'asmaa-abdelrahim': 'asmaa_abdelrahim', 'mahmoud-elsadd': 'mahmoud_elsadd',
  'gomaa-fahmy': 'gomaa_fahmy', 'eman-awad': 'eman_awad', 'tamer-elsharkawy': 'tamer_elsharkawy_TA',
};

const masters = [];
for (const name of readdirSync(join(samples, 'masters'))) {
  masters.push(await readMaster({ programme: name.split('_')[0], name, data: readFileSync(join(samples, 'masters', name)) }));
}
const result = generate(masters, SEED_ROSTER);
const diffs = new Map<string, Awaited<ReturnType<typeof compare>>>();
for (const [id, file] of Object.entries(FILES)) {
  const expected = await readExpected(readFileSync(join(samples, 'expected', `${file}.xlsx`)));
  diffs.set(id, compare(result.sheets.find((s) => s.staff.id === id)!, expected));
}

test('diff against every hand-made sheet (printed, informational)', () => {
  for (const [id, d] of diffs) {
    const l = d.lectureSlots;
    const s = d.supervisionSlots;
    console.log(`\n${id}: teaching slots ${l.same}/${l.expected} of the sheet (generated ${l.generated}); supervision ${s.same}/${s.expected} (generated ${s.generated})`);
    for (const line of d.lines) console.log(`  - ${line}`);
  }
  assert.equal(diffs.size, 10);
});

test('Asmaa Radi: every lecture slot of her sheet that comes from an available master is generated', () => {
  const d = diffs.get('asmaa-radi')!;
  // 12 lecture hours on her sheet, 2 of them from the mechanics master (PDF only).
  assert.deepEqual([d.lectureSlots.same, d.lectureSlots.expected], [10, 12]);
  assert.deepEqual(d.lines.filter((x) => x.startsWith('course differs')), []);
});

test('most lecture slots of the comms staff match their sheets', () => {
  for (const id of ['asmaa-radi', 'ahmed-salem', 'hend-elsayed', 'mai-helmy']) {
    const l = diffs.get(id)!.lectureSlots;
    assert.ok(l.same / l.expected >= 0.7, `${id}: ${l.same}/${l.expected}`);
  }
});
