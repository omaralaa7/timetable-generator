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

test('Asmaa Radi: courses, shares and weeks as on her sheet (the mechanics course is remembered from her sheet)', () => {
  assert.deepEqual(sheet('asmaa-radi').courses.map((c) => `${c.name} | ${c.share} | ${c.weeks}`), [
    'مبادئ بث المعلومات (ELC 151) | 0.5 | 1 - 7',
    'رياضيات 3 (MAT 101) | 0.5 | 9 - 15',
    'الكترونيات الجوامد والنبائط (ELC 211) | 0.5 | 9 - 15',
    'الاتصالات التناظرية (ELC 351) | 0.5 | 9 - 15',
    'الاتصالات الرقمية (ELC 451) | 0.5 | 1 - 7',
    'دوائر كهربية والكترونية (ELP 331 ) | 0.5 | 1 - 7',
  ]);
  // as on her sheet (this test loads comms + power only; the civil master adds a MAT 101 lecture)
  assert.equal(sheet('asmaa-radi').totals.teaching, 12);
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
    for (const it of s.items.filter((x) => x.type === 'supervision' && x.sessionId)) {
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

test('PDF wrapper: valid structure, one page per picture', async () => {
  const { imagesToPdf } = await import('../src/core/pdf.ts');
  const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3, 0xff, 0xd9]);
  const pdf = imagesToPdf([{ jpeg, width: 20, height: 10 }, { jpeg, width: 20, height: 10 }]);
  const text = Buffer.from(pdf).toString('latin1');
  assert.ok(text.startsWith('%PDF-1.4'));
  assert.ok(text.trimEnd().endsWith('%%EOF'));
  assert.match(text, /\/Count 2/);
  assert.equal(text.match(/\/Filter \/DCTDecode/g)?.length, 2);
  // every cross-reference entry points at the start of its object
  const xref = Number(/startxref\n(\d+)/.exec(text)![1]);
  assert.ok(text.startsWith('xref', xref));
  const entries = text.slice(xref).split('\n').slice(3).filter((l) => / n $/.test(l));
  entries.forEach((line, i) => assert.ok(text.startsWith(`${i + 1} 0 obj`, Number(line.slice(0, 10))), `object ${i + 1}`));
  assert.equal(entries.length, 8);
});

test('a lecture written once with its code and once without is one course; the code-less copy is not counted', () => {
  const rows = sheet('asmaa-abdelrahim').courses.filter((c) => c.name.includes('نظم القياسات'));
  assert.deepEqual(rows.map((c) => `${c.name} | ${c.share} | ${c.weeks}`), ['اختيارى 1 (ELP 3E1) (نظم القياسات) | 0.5 | 9 - 15']);
  const cells = sheet('asmaa-abdelrahim').items.filter((it) => it.type === 'lecture' && it.text.includes('نظم القياسات'));
  // one lecture a week (Monday 02:15 in the corrected power master), written with its code
  assert.deepEqual(cells.map((it) => `${it.day}:${it.slots.join(',')}`), ['2:6,7']);
  assert.ok(cells[0].text.includes('(ELP 3E1)'));
});

test('duties follow the client notes: fixed office hours, none on Saturday, TA without office or quality', () => {
  for (const s of result.sheets) {
    const saturday = s.items.filter((it) => it.day === 0 && (it.type === 'office' || it.type === 'advising' || it.type === 'quality'));
    assert.deepEqual(saturday.map((it) => it.text), [], s.staff.name);
    assert.ok(s.totals.office <= (s.staff.officeHours ?? 4), `${s.staff.name}: office ${s.totals.office}`);
  }
  assert.equal(sheet('gomaa-fahmy').totals.office, 6);
  assert.equal(sheet('mahmoud-elsadd').totals.office, 6);
  for (const s of result.sheets) {
    // at most 4, except the two the client allowed more so that they reach their minimum
    const cap = { 'mai-helmy': 6, 'tamer-elsharkawy': 8 }[s.staff.id] ?? 4;
    assert.ok(s.totals.advising <= cap, `${s.staff.name}: advising ${s.totals.advising}`);
    assert.ok(s.totals.quality <= 2, `${s.staff.name}: quality ${s.totals.quality}`);
  }
  assert.equal(sheet('asmaa-radi').totals.office, 4);
  for (const id of ['ahmed-salem', 'mahmoud-elsadd', 'gomaa-fahmy']) assert.equal(sheet(id).totals.advising, 0, id);
  const tamer = sheet('tamer-elsharkawy').totals;
  assert.deepEqual([tamer.office, tamer.quality], [0, 0]);
});

test('Tuesday third period (12:30–02:10) is kept for activities on every sheet', () => {
  for (const s of result.sheets) {
    const there = s.items.filter((it) => it.day === 3 && it.slots.some((x) => x === 4 || x === 5));
    assert.deepEqual(there.map((it) => `${it.type}:${it.slots.join(',')}`), ['activity:4,5'], s.staff.name);
    // not counted: the day's total only has the other blocks
    const others = new Set(s.items.filter((it) => it.day === 3 && it.type !== 'activity').flatMap((it) => it.slots)).size;
    assert.equal(s.daily[3], others, s.staff.name);
  }
});
