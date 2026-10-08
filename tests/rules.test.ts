import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { test } from 'node:test';
import ExcelJS from 'exceljs';
import { generate, readMaster } from '../src/core/pipeline.ts';
import { SEED_ROSTER } from '../src/core/roster.ts';
import { writeSheet } from '../src/core/writer.ts';

const dir = join(import.meta.dirname, '..', 'samples', 'masters');
const masters = [];
for (const programme of ['comms', 'power']) {
  const name = `${programme}_2026-27.docx`;
  masters.push(await readMaster({ programme, name, data: readFileSync(join(dir, name)) }));
}
const result = generate(masters, SEED_ROSTER);
const sheet = (id: string) => result.sheets.find((s) => s.staff.id === id)!;

test('Asmaa Radi: courses, shares and weeks as on her sheet (mechanics course excluded — PDF only)', () => {
  assert.deepEqual(sheet('asmaa-radi').courses.map((c) => `${c.name} | ${c.share} | ${c.weeks}`), [
    'مبادئ بث المعلومات (ELC 151) | 0.5 | 1 - 7',
    'رياضيات 3 (MAT 101) | 0.5 | 9 - 15',
    'الكترونيات الجوامد والنبائط (ELC 211) | 0.5 | 9 - 15',
    'الاتصالات التناظرية (ELC 351) | 0.5 | 9 - 15',
    'الاتصالات الرقمية (ELC 451) | 0.5 | 1 - 7',
  ]);
  assert.equal(sheet('asmaa-radi').totals.teaching, 10);
});

test('nobody is double-booked or placed on their secondment day by the rules', () => {
  for (const s of result.sheets) {
    const auto = s.items.filter((it) => it.type === 'supervision' || it.origin === 'auto');
    for (const it of auto) {
      assert.notEqual(it.day, s.staff.secondmentDay, `${s.staff.name}: ${it.text}`);
      const clash = s.items.find((o) => o !== it && o.day === it.day && o.slots.some((x) => it.slots.includes(x)));
      assert.equal(clash, undefined, `${s.staff.name}: ${it.text} × ${clash?.text}`);
    }
  }
});

test('a section is supervised by exactly one of its listed doctors', () => {
  const seen = new Map<string, string>();
  for (const s of result.sheets) {
    for (const it of s.items.filter((x) => x.type === 'supervision')) {
      assert.equal(seen.has(it.sessionId!), false, `${it.text} given to ${seen.get(it.sessionId!)} and ${s.staff.name}`);
      seen.set(it.sessionId!, s.staff.name);
    }
  }
});

test('one-doctor lecture: share 1, weeks 1–15; the TA gets sections, not supervision', () => {
  const control = sheet('ahmed-elsayed').courses.find((c) => c.name.includes('التحكم وتطبيقات الحاسوب'))!;
  assert.deepEqual([control.share, control.weeks], [1, '1 - 15']);
  const tamer = sheet('tamer-elsharkawy');
  assert.equal(tamer.totals.supervision, 0);
  assert.ok(tamer.totals.teaching > 15);
});

test('written sheet: layout anchors of the sample', async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await writeSheet(sheet('asmaa-radi')) as unknown as ArrayBuffer);
  const ws = wb.worksheets[0];
  const n = sheet('asmaa-radi').courses.length;
  assert.equal(ws.getCell('D3').value, 'جدول الأعباء التدريسية وملحقاتها');
  assert.equal(ws.getCell('B7').value, 'الإسم: د/أسماء راضي');
  assert.equal(ws.getCell(`C${12 + n}`).value, '09:00 - 09:50');
  assert.equal(ws.getCell(`B${13 + n}`).value, 'السبت');
  assert.match(String(ws.getCell(`C${13 + n}`).value), /^انت/);
  assert.equal(ws.getCell(`B${22 + n}`).value, 'النصاب');
  assert.equal(ws.views[0].rightToLeft, true);
  assert.equal(ws.pageSetup.paperSize, 9);
});
