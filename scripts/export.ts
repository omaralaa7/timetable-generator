// Write every roster member's sheet from the sample masters into out/.
//   node scripts/export.ts
import { mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { generate, readMaster } from '../src/core/pipeline.ts';
import { SEED_ROSTER } from '../src/core/roster.ts';
import { sheetFileName, writeSheet } from '../src/core/writer.ts';

const root = join(import.meta.dirname, '..');
const dir = join(root, 'samples', 'masters');
const masters = [];
for (const name of readdirSync(dir).filter((f) => f.endsWith('.docx') || f.endsWith('.xlsx'))) {
  masters.push(await readMaster({ programme: name.split('_')[0], name, data: readFileSync(join(dir, name)) }));
}
const { sheets } = generate(masters, SEED_ROSTER);
mkdirSync(join(root, 'out'), { recursive: true });
for (const sheet of sheets) {
  writeFileSync(join(root, 'out', sheetFileName(sheet)), await writeSheet(sheet));
  console.log('wrote', sheetFileName(sheet));
}
