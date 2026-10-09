import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { test } from 'node:test';
import { parseCellText } from '../src/core/cellText.ts';
import { readDocx } from '../src/core/docx.ts';
import { matchStaff, SEED_ROSTER } from '../src/core/roster.ts';
import { parseMaster } from '../src/core/sessions.ts';
import { parseSlot } from '../src/core/slots.ts';
import type { ParsedMaster, Session } from '../src/core/types.ts';

const masters = join(import.meta.dirname, '..', 'samples', 'masters');

async function load(name: string): Promise<ParsedMaster> {
  const master = parseMaster(await readDocx(readFileSync(join(masters, `${name}_2026-27.docx`))), name);
  matchStaff(master.sessions, SEED_ROSTER);
  return master;
}

const comms = await load('comms');
const power = await load('power');

function of(staffId: string, type: string, ...lists: ParsedMaster[]): Session[] {
  return lists
    .flatMap((m) => m.sessions)
    .filter((s) => s.parsed.type === type && s.parsed.names.some((n) => n.staffId === staffId));
}

const at = (s: Session) => `${s.programme} ${s.day}:${s.slots.join(',')} L${s.level} ${s.wholeLevel ? 'all' : s.sections.join('+')}`;

test('time labels map to slots by start time, typos tolerated', () => {
  assert.deepEqual(parseSlot('09:00 – 09:50'), { slot: 0, exact: true });
  assert.deepEqual(parseSlot('09:00 09:50'), { slot: 0, exact: true });
  assert.deepEqual(parseSlot('12:30 – 01:20'), { slot: 4, exact: true });
  assert.equal(parseSlot('5:40 – 4:50')?.slot, 9);
  assert.equal(parseSlot('09.40 - 10.40')?.slot, 1);
  assert.equal(parseSlot('08:50')?.slot, 0);
  assert.equal(parseSlot('المستوي'), null);
});

test('cell text: lecture with two doctors, one underlined', () => {
  const p = parseCellText([
    { text: 'محــاضـــرة مبـــادئ بــث المعلـومــــات (ELC 151) (د. أحمد سالم & د. ', underline: false },
    { text: 'اسماء راضى', underline: true },
    { text: ') / مدرج 7 (No. 722)', underline: false },
  ]);
  assert.equal(p.type, 'lecture');
  assert.equal(p.title, 'مبادئ بث المعلومات');
  assert.equal(p.code, 'ELC 151');
  assert.equal(p.room, 'مدرج 7 (No. 722)');
  assert.deepEqual(p.names.map((n) => [n.name, n.underline > 0.5]), [['أحمد سالم', false], ['اسماء راضى', true]]);
});

test('cell text: quirks from the masters', () => {
  const parse = (text: string) => parseCellText([{ text, underline: false }]);
  // reversed bracket, `&` inside the room
  let p = parse('ت عملى مبادئ بث المعلومات)م. شهد) (د. أحمد سالم& د اسماء راضى ) / Lab 714 & قاعة 702');
  assert.equal(p.type, 'section');
  assert.equal(p.title, 'مبادئ بث المعلومات');
  assert.equal(p.room, 'Lab 714 & قاعة 702');
  assert.deepEqual(p.names.map((n) => `${n.role}:${n.name}`), ['ta:شهد', 'doctor:أحمد سالم', 'doctor:اسماء راضى']);
  // missing closing bracket, room starting with `م` (= معمل)
  p = parse('ت عملى  تحليل شبكات القوى (ا.د. محمود السد & ا.م.د. جمعة فهمى )(م. تامر م خطوط النقل');
  assert.deepEqual(p.names.map((n) => `${n.rank}:${n.name}`), ['professor:محمود السد', 'associate:جمعة فهمى', 'ta:تامر']);
  assert.equal(p.room, 'م خطوط النقل');
  // elective without `محاضرة`, code in square brackets
  p = parse('اختيارى 1 (نظم القياسات) [ELP3E1] (د. اسماء عبد الرحيم & د. سيد محمد احمد )  قاعة  رقم 431');
  assert.deepEqual([p.type, p.title, p.subtitle, p.code, p.room], ['lecture', 'اختيارى 1', 'نظم القياسات', 'ELP 3E1', 'قاعة رقم 431']);
  // code without brackets / number after the code / missing second name
  assert.deepEqual(
    (({ title, code }) => [title, code])(parse('محــاضـــرة مبـــادئ بــث المعلـومــــات ELC 151  (د. أحمد سالم) / مدرج 7')),
    ['مبادئ بث المعلومات', 'ELC 151'],
  );
  assert.equal(parse('محاضرة رياضيات  (MAT 101)3 ( د.أسماء عبدالرحيم & د. اسماء راضى) مدرج 7').title, 'رياضيات 3');
  assert.deepEqual(parse('محاضرة س (ELP 111) (د. أحمد السيد&) مدرج 7').names.map((n) => n.name), ['أحمد السيد']);
  assert.equal(parse('ت عملى مشروع التخرج').type, 'project');
});

test('both Word masters parse with the expected structure', () => {
  for (const m of [comms, power]) {
    assert.ok(m.sessions.length > 80, `${m.programme}: ${m.sessions.length} sessions`);
    assert.deepEqual([...new Set(m.sessions.map((s) => s.level))].sort(), [1, 2, 3, 4]);
    // power has nothing on Thursday
    const days = m === comms ? [0, 1, 2, 3, 4, 5] : [0, 1, 2, 3, 4];
    assert.deepEqual([...new Set(m.sessions.map((s) => s.day))].sort(), days);
    for (const s of m.sessions) {
      assert.ok(s.sections.every((x) => x === 1 || x === 2), at(s));
    }
  }
  assert.equal(comms.warnings.length, 0);
  // power has one orphan cell (`/ Lab 825`) left behind in the master
  assert.deepEqual(power.warnings.map((w) => w.kind), ['unknown-type']);
});

test('Asmaa Radi: lectures and who is underlined match her sheet', () => {
  const lectures = of('asmaa-radi', 'lecture', comms, power).map((s) => {
    const me = s.parsed.names.find((n) => n.staffId === 'asmaa-radi')!;
    const other = s.parsed.names.find((n) => n !== me)!;
    const weeks = me.underline > 0.5 ? '1-7' : other.underline > 0.5 ? '9-15' : '?';
    return `${at(s)} ${s.parsed.code} ${weeks}`;
  });
  assert.deepEqual(lectures.sort(), [
    'comms 1:0,1 L1 all ELC 151 1-7',
    'comms 1:2,3 L4 all ELC 451 1-7',
    'comms 2:0,1 L2 all ELC 211 9-15',
    'comms 2:2,3 L3 all ELC 351 9-15',
    'comms 3:0,1 L1 all MAT 101 9-15',
    'power 1:0,1 L1 all ELC 151 1-7',
    'power 3:0,1 L1 all MAT 101 9-15',
  ]);
});

test('Asmaa Radi: every supervision on her sheet is among her section sessions', () => {
  const mine = of('asmaa-radi', 'section', comms, power);
  const has = (day: number, slot: number, title: string) =>
    mine.some((s) => s.day === day && s.slots.includes(slot) && s.parsed.title === title);
  // (day, slot, course) from samples/expected/asmaa_radi.xlsx
  const sheet: [number, number, string][] = [
    [1, 4, 'الاتصالات الرقمية'], [1, 6, 'رياضيات 3'], [1, 8, 'مبادئ بث المعلومات'],
    [2, 4, 'الاتصالات التناظرية'], [2, 9, 'الاتصالات التناظرية'],
    [3, 8, 'رياضيات 3'], [4, 2, 'مبادئ بث المعلومات'], [4, 4, 'الكترونيات الجوامد والنبائط'],
  ];
  for (const [day, slot, title] of sheet) assert.ok(has(day, slot, title), `missing ${day}:${slot} ${title}`);
  assert.equal(mine.length, 17);
});

test('roster: similar names stay apart, one-word names are not guessed', () => {
  const all = [...comms.sessions, ...power.sessions];
  const ids = (written: string) =>
    new Set(all.flatMap((s) => s.parsed.names).filter((n) => n.name === written).map((n) => n.staffId));
  assert.deepEqual([...ids('اسماء راضى')], ['asmaa-radi']);
  assert.deepEqual([...ids('أسماء عبدالرحيم')], ['asmaa-abdelrahim']);
  assert.deepEqual([...ids('أحمد سالم')], ['ahmed-salem']);
  assert.deepEqual([...ids('أحمد السيد')], ['ahmed-elsayed']);
  // جمعة عثمان is the same person as جمعه فهمى (full name جمعة فهمي عثمان)
  assert.deepEqual([...ids('جمعة عثمان')], ['gomaa-fahmy']);
  assert.deepEqual([...ids('مايسة')], [null]);
  assert.equal(of('tamer-elsharkawy', 'section', power).length > 10, true);
});
