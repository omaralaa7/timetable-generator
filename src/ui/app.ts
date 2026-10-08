import JSZip from 'jszip';
import { generate, readMaster } from '../core/pipeline.ts';
import type { Staff } from '../core/roster.ts';
import { type GridItem, type Issue, type ItemType, type Result, type StaffSheet } from '../core/rules.ts';
import { SLOT_LABELS } from '../core/slots.ts';
import type { Rank } from '../core/types.ts';
import { sheetFileName, writeSheet } from '../core/writer.ts';
import { loadProject, saveProject, type Project } from './state.ts';

type Tab = 'files' | 'sheets' | 'issues' | 'settings';

const DAY_NAMES = ['السبت', 'الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس'];
const RANKS: Record<Rank, string> = { professor: 'أستاذ دكتور', associate: 'أستاذ مساعد', lecturer: 'مدرس', ta: 'مدرس مساعد / معيد' };
const TYPES: Record<ItemType, string> = {
  lecture: 'محاضرة', supervision: 'إشراف', section: 'تمرين / عملي',
  office: 'ساعات مكتبية', advising: 'ارشاد أكاديمي', quality: 'جودة',
};
const DUTY_TEXT: Partial<Record<ItemType, string>> = { office: 'ساعات مكتبية', advising: 'ارشاد أكاديمي', quality: 'جودة' };
const ISSUE_TITLES: Record<Issue['kind'], string> = {
  'no-underline': 'محاضرات بدون خط تحت اسم من يبدأ',
  'short-name': 'أسماء مختصرة',
  conflict: 'حصص لا يوجد لها مشرف متاح',
  'both-free': 'حصص الدكتوران متاحان فيها (اقتراح)',
  overlap: 'تعارضات في الجدول الرئيسي',
  'below-minimum': 'أقل من النصاب',
  project: 'مشروع التخرج',
};

const project: Project = loadProject();
let tab: Tab = project.masters.length ? 'sheets' : 'files';
let person = '';
let message = '';
let result: Result = generate(project.masters, project.roster, project.decisions);
/** Uploaded Excel files kept in memory so another sheet can be chosen without re-uploading. */
const fileData = new Map<string, ArrayBuffer>();
const root = document.getElementById('app')!;

const esc = (s: unknown) =>
  String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);

function commit(note = ''): void {
  result = generate(project.masters, project.roster, project.decisions);
  if (!project.masters.length && (tab === 'sheets' || tab === 'issues')) tab = 'files';
  message = saveProject(project) ? note : 'تعذر الحفظ في المتصفح — ستبقى التعديلات حتى إغلاق الصفحة فقط';
  render();
}

function download(name: string, data: BlobPart, type: string): void {
  const url = URL.createObjectURL(new Blob([data], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ───────────────────────── views ─────────────────────────

function render(): void {
  const ready = project.masters.length > 0;
  const open = result.issues.length;
  const tabs: [Tab, string, boolean][] = [
    ['files', '١ · رفع الجداول الرئيسية', true],
    ['sheets', '٢ · الجداول الفردية', ready],
    ['issues', `٣ · قرارات اختيارية${ready && open ? ` (${open})` : ''}`, ready],
    ['settings', '⚙ الإعدادات', true],
  ];
  root.innerHTML = `
    <header><h1>مولّد الجداول الفردية</h1></header>
    <nav>${tabs.map(([id, label, on]) => `<button data-tab="${id}" class="${id === tab ? 'on' : ''}" ${on ? '' : 'disabled'}>${label}</button>`).join('')}</nav>
    ${message ? `<p class="note">${esc(message)}</p>` : ''}
    <main>${{ files: filesView, sheets: sheetsView, issues: issuesView, settings: settingsView }[tab]()}</main>
    <dialog id="editor"></dialog>`;
}

function filesView(): string {
  const masters = project.masters.map((m, i) => `
    <li>
      <div class="row">
        <span class="ok">✓</span>
        <strong>${esc(m.fileName)}</strong>
        <span class="hint">${m.sessions.length} حصة</span>
        ${m.sheets && m.sheets.length > 1 ? `<label>الورقة:
          <select data-act="sheet" data-i="${i}" ${fileData.has(m.fileName) ? '' : 'disabled title="أعد رفع الملف لتغيير الورقة"'}>
            ${m.sheets.map((s) => `<option ${s === m.sheet ? 'selected' : ''}>${esc(s)}</option>`).join('')}
          </select></label>` : ''}
        <button data-act="remove-master" data-i="${i}" class="link danger">حذف</button>
      </div>
      ${m.warnings.length ? `<details><summary class="small">${m.warnings.length} خانة لم تُفهم (للاطلاع فقط)</summary><ul>${m.warnings.map((w) => `<li>${esc(w.message)}</li>`).join('')}</ul></details>` : ''}
    </li>`).join('');
  return `
    <section>
      <p class="lead">الخطوة ١: ضع ملفات الجداول الرئيسية لكل البرامج (اتصالات، قوى، مدني…).</p>
      <label class="drop" id="drop">
        <input type="file" multiple accept=".docx,.xlsx,.pdf" data-act="upload" hidden>
        <b>اسحب الملفات إلى هنا أو اضغط للاختيار</b>
        <span>ملفات Word أو Excel — الملفات لا تغادر جهازك</span>
      </label>
      <ul class="masters">${masters}</ul>
      <div class="next">
        <button data-act="process" class="primary big" ${project.masters.length ? '' : 'disabled'}>إنشاء الجداول الفردية ←</button>
        ${project.masters.length ? '' : '<span class="hint">ارفع ملفاً واحداً على الأقل أولاً</span>'}
      </div>
      <details class="guide">
        <summary>طريقة الاستخدام</summary>
        <ol>
          <li><b>ارفع الجداول الرئيسية</b> ثم اضغط «إنشاء الجداول الفردية».</li>
          <li><b>الجداول الفردية</b>: اختر العضو لترى جدوله، واضغط «تنزيل كل الجداول» للحصول على ملفات Excel.</li>
          <li><b>قرارات اختيارية</b>: حالات لم يستطع البرنامج حسمها وحده (مثلاً من يشرف على معمل). إن تركتها يُستخدم اقتراح البرنامج.</li>
        </ol>
        <p>إذا تغيّر جدول رئيسي، ارفعه مرة أخرى بنفس اسم الملف وسيُستبدل القديم وتبقى قراراتك. ملفات PDF غير مدعومة.</p>
        <p>القاعدة: الدكتور الذي تحت اسمه خط في خانة المحاضرة يدرّس الأسابيع 1–7 والآخر 9–15.</p>
      </details>
    </section>`;
}

function issueCard(i: Issue): string {
  const buttons = (i.options ?? []).map((o, n) => {
    const label = i.labels?.[o] ?? o;
    const text = i.kind === 'no-underline' ? `يبدأ: ${label}` : i.kind === 'both-free' && n === 0 ? `✓ ${label}` : label;
    return `<button data-act="fix" data-id="${esc(i.id)}" data-opt="${esc(o)}" class="${i.kind === 'both-free' && n === 0 ? 'primary' : ''}">${esc(text)}</button>`;
  }).join('');
  const none = i.kind === 'conflict' ? `<button data-act="fix" data-id="${esc(i.id)}" data-opt="none">لا أحد من القسم</button>` : '';
  const go = i.staffIds.length ? `<button data-act="person" data-id="${esc(i.staffIds[0])}" class="link">عرض الجدول</button>` : '';
  return `<li><p>${esc(i.message)}</p><div class="row">${buttons}${none}${go}<button data-act="dismiss" data-id="${esc(i.id)}" class="link">تجاهل</button></div></li>`;
}

function gridTable(sheet: StaffSheet): string {
  const rows = DAY_NAMES.map((name, day) => {
    if (sheet.staff.secondmentDay === day) return `<tr><th>${name}</th><td colspan="10" class="secondment">انتداب</td><td></td></tr>`;
    let cells = '';
    const taken = new Set<number>();
    for (let slot = 0; slot < 10; slot++) {
      if (taken.has(slot)) continue;
      const item = sheet.items.find((it) => it.day === day && it.slots.includes(slot));
      if (!item) {
        cells += `<td data-cell="${day}:${slot}" class="free"></td>`;
        continue;
      }
      let span = 0;
      while (item.slots.includes(slot + span) && !taken.has(slot + span)) taken.add(slot + span++);
      cells += `<td data-cell="${day}:${slot}" colspan="${span}" class="t-${item.type}">${esc(item.text)}</td>`;
    }
    return `<tr><th>${name}</th>${cells}<td class="total">${sheet.daily[day] || ''}</td></tr>`;
  }).join('');
  return `<div class="scroll"><table class="grid">
    <thead><tr><th></th>${SLOT_LABELS.map((l) => `<th>${l}</th>`).join('')}<th>المجموع</th></tr></thead>
    <tbody>${rows}</tbody></table></div>`;
}

function personView(sheet: StaffSheet): string {
  const t = sheet.totals;
  const isTa = sheet.staff.rank === 'ta';
  const edited = !!(project.decisions.edits?.[sheet.staff.id] || project.decisions.duties?.[sheet.staff.id] || project.decisions.courses?.[sheet.staff.id]);
  const courses = sheet.courses.map((c, i) => `
    <tr data-i="${i}">
      <td><input data-c="name" value="${esc(c.name)}"></td>
      <td><input data-c="share" type="number" min="0" max="1" step="0.05" value="${c.share}"></td>
      <td><input data-c="weeks" value="${esc(c.weeks)}"></td>
      <td><button data-act="remove-course" data-i="${i}" class="danger">✕</button></td>
    </tr>`).join('');
  return `
    <div class="person">
      <div class="row">
        <h3>${esc(sheet.staff.name)} — ${RANKS[sheet.staff.rank]}</h3>
        <button data-act="download-one" data-id="${esc(sheet.staff.id)}">تنزيل هذا الجدول (Excel)</button>
        ${edited ? `<button data-act="reset-person" data-id="${esc(sheet.staff.id)}" class="link">إلغاء تعديلاتي</button>` : ''}
      </div>
      ${gridTable(sheet)}
      <p class="hint"><span class="key t-lecture">محاضرة</span> <span class="key t-supervision">إشراف</span> <span class="key t-office">مكتبية / ارشاد / جودة</span> — للتعديل اضغط على أي خانة.</p>
      <table class="totals">
        <tr><th>${isTa ? 'تمرين/عملي' : 'المحاضرات'}</th><th>الساعات المكتبية</th><th>ساعات الإشراف</th><th>الإرشاد الأكاديمي</th><th>الجودة</th><th>النصاب (بدون الجودة)</th></tr>
        <tr><td>${t.teaching}</td><td>${t.office}</td><td>${t.supervision}</td><td>${t.advising}</td><td>${t.quality}</td>
          <td class="${t.counted < t.minimum ? 'bad' : 'good'}">${t.counted} من ${t.minimum}</td></tr>
      </table>
      <details>
        <summary>المقررات المشارك فيها (${sheet.courses.length})</summary>
        <table class="courses">
          <thead><tr><th>المقرر</th><th>نسبة المشاركة</th><th>الأسابيع</th><th></th></tr></thead>
          <tbody>${courses}</tbody>
        </table>
        <button data-act="add-course">+ إضافة مقرر</button>
      </details>
    </div>`;
}

function sheetsView(): string {
  const current = result.sheets.find((s) => s.staff.id === person) ?? result.sheets[0];
  const open = result.issues.length;
  return `
    <section>
      <div class="done">
        <div>
          <b>تم إنشاء ${result.sheets.length} جدولاً فردياً.</b>
          ${open ? `<span class="hint">يوجد ${open} قرار اختياري — <button data-tab="issues" class="link">عرضها</button></span>` : ''}
        </div>
        <button data-act="download-all" class="primary big">⬇ تنزيل كل الجداول (ملف مضغوط)</button>
      </div>
      <p class="lead">اختر عضواً لمعاينة جدوله:</p>
      <div class="people">${result.sheets.map((s) => `<button data-act="person" data-id="${esc(s.staff.id)}" class="${s === current ? 'on' : ''} ${s.totals.counted < s.totals.minimum ? 'low' : ''}">${esc(s.staff.name)}</button>`).join('')}</div>
      ${current ? personView(current) : '<p class="empty">لا يوجد أعضاء في القائمة — أضفهم من الإعدادات.</p>'}
    </section>`;
}

function issuesView(): string {
  const kinds = Object.keys(ISSUE_TITLES) as Issue['kind'][];
  const groups = kinds.map((kind) => {
    const list = result.issues.filter((i) => i.kind === kind);
    if (!list.length) return '';
    return `<details><summary>${ISSUE_TITLES[kind]} <b>${list.length}</b></summary><ul class="issues">${list.map(issueCard).join('')}</ul></details>`;
  }).join('');
  const decided = Object.keys(project.decisions.supervisor ?? {}).length + Object.keys(project.decisions.starts ?? {}).length +
    Object.keys(project.decisions.dismissed ?? {}).length + Object.keys(project.decisions.names ?? {}).length;
  return `
    <section>
      <p class="lead">هذه الخطوة اختيارية. هذه حالات لم يستطع البرنامج حسمها وحده؛ إن تركتها يبقى اقتراحه كما هو. اضغط على أي مجموعة لفتحها.</p>
      ${groups || '<p class="good">لا توجد قرارات معلّقة.</p>'}
      ${decided ? `<button data-act="undo-decisions" class="link">التراجع عن قراراتي (${decided})</button>` : ''}
      <div class="next"><button data-tab="sheets" class="primary">← العودة إلى الجداول</button></div>
    </section>`;
}

function settingsView(): string {
  const s = project.settings;
  const dayOptions = (value: number | null) =>
    `<option value="">—</option>${DAY_NAMES.map((d, i) => `<option value="${i}" ${value === i ? 'selected' : ''}>${d}</option>`).join('')}`;
  const rows = project.roster.map((st, i) => `
    <tr data-i="${i}">
      <td><input data-f="name" value="${esc(st.name)}"></td>
      <td><input data-f="aliases" value="${esc(st.aliases.join('، '))}" title="الأسماء كما تُكتب في الجداول الرئيسية، مفصولة بفاصلة"></td>
      <td><input data-f="programme" value="${esc(st.programme)}"></td>
      <td><select data-f="rank">${(Object.keys(RANKS) as Rank[]).map((r) => `<option value="${r}" ${st.rank === r ? 'selected' : ''}>${RANKS[r]}</option>`).join('')}</select></td>
      <td class="c"><input type="checkbox" data-f="advising" ${st.advising ? 'checked' : ''}></td>
      <td><select data-f="secondmentDay">${dayOptions(st.secondmentDay)}</select></td>
      <td><input type="number" min="0" max="10" data-f="qualityHours" value="${st.qualityHours}"></td>
      <td><input type="number" min="0" max="20" data-f="officeHours" value="${st.officeHours ?? 4}"></td>
      <td><input type="number" min="0" max="20" data-f="advisingHours" value="${st.advisingHours ?? 4}" ${st.advising ? '' : 'disabled'}></td>
      <td><button data-act="remove-staff" data-i="${i}" class="danger" title="حذف">✕</button></td>
    </tr>`).join('');
  const detected = result.detected.filter((d) => d.role === 'doctor' || d.count > 2);
  return `
    <section>
      <p class="lead">لا تحتاج هذه الصفحة في الاستخدام العادي. استخدمها عند تغيّر أعضاء القسم أو بيانات التوقيعات.</p>
      <details open>
        <summary>أعضاء القسم (${project.roster.length})</summary>
        <p class="hint">فقط من في هذه القائمة يصدر له جدول.</p>
        <div class="scroll"><table class="roster">
          <thead><tr><th>الاسم على الجدول</th><th>الاسم كما يُكتب في الجداول الرئيسية</th><th>البرنامج</th><th>الدرجة</th><th>مرشد</th><th>الانتداب</th><th>جودة</th><th>مكتبية</th><th>ارشاد</th><th></th></tr></thead>
          <tbody>${rows}</tbody>
        </table></div>
        <button data-act="add-staff">+ إضافة عضو</button>
        ${detected.length ? `<details><summary class="small">أسماء في الجداول الرئيسية ليست في القائمة (${detected.length})</summary>
          <p class="hint">أضف من هو من القسم فقط؛ الباقون يُتجاهلون.</p>
          <ul class="chips">${detected.map((d) => `<li>${esc(d.name)} <button data-act="add-detected" data-name="${esc(d.name)}" data-role="${d.role}">إضافة</button></li>`).join('')}</ul></details>` : ''}
      </details>
      <details>
        <summary>بيانات تظهر في كل جدول (الفصل الدراسي والتوقيعات)</summary>
        <div class="form">
          <label>الفصل الدراسي <input data-s="semester" value="${esc(s.semester)}"></label>
          <label>اسم القسم <input data-s="department" value="${esc(s.department)}"></label>
          ${s.signatures.map((sig, i) => `<label>${esc(sig.title)} <input data-sig="${i}" value="${esc(sig.name)}"></label>`).join('')}
        </div>
      </details>
    </section>`;
}

// ───────────────────────── actions ─────────────────────────

async function addFiles(files: File[]): Promise<void> {
  const notes: string[] = [];
  for (const file of files) {
    try {
      const data = await file.arrayBuffer();
      const programme = file.name.replace(/\.[^.]+$/, '');
      const master = await readMaster({ programme, name: file.name, data });
      if (file.name.toLowerCase().endsWith('.xlsx')) fileData.set(file.name, data);
      const stored = { programme, fileName: file.name, sheet: master.sheet, sheets: master.sheets, sessions: master.sessions, warnings: master.warnings };
      const at = project.masters.findIndex((m) => m.fileName === file.name);
      if (at >= 0) project.masters[at] = stored;
      else project.masters.push(stored);
      if (at >= 0) notes.push(`«${file.name}»: تم استبدال النسخة السابقة`);
    } catch (e) {
      const text = e instanceof Error ? e.message : String(e);
      // Library errors are in English; show our own wording instead.
      notes.push(/[؀-ۿ]/.test(text) ? text : `«${file.name}»: تعذر قراءة الملف — تأكد أنه ملف Word أو Excel سليم`);
    }
  }
  commit(notes.join(' — '));
}

function staffSheet(id: string): StaffSheet | undefined {
  return result.sheets.find((s) => s.staff.id === id);
}

/** Before the first manual change, freeze the suggested duties so they stop moving around. */
function freezeDuties(sheet: StaffSheet): void {
  const d = project.decisions;
  d.duties ??= {};
  if (d.duties[sheet.staff.id]) return;
  d.duties[sheet.staff.id] = sheet.items
    .filter((it) => it.origin !== 'master')
    .map((it): GridItem => ({ day: it.day, slots: [...it.slots], type: it.type, text: it.text, origin: 'manual' }));
}

function openEditor(cell: string): void {
  const sheet = staffSheet(person) ?? result.sheets[0];
  if (!sheet) return;
  const [day, slot] = cell.split(':').map(Number);
  const item = sheet.items.find((it) => it.day === day && it.slots.includes(slot));
  const types: (ItemType | 'empty')[] = sheet.staff.rank === 'ta'
    ? ['section', 'office', 'advising', 'quality', 'empty']
    : ['lecture', 'supervision', 'office', 'advising', 'quality', 'empty'];
  const dialog = document.getElementById('editor') as HTMLDialogElement;
  dialog.innerHTML = `
    <form method="dialog">
      <h3>${DAY_NAMES[day]} — ${SLOT_LABELS[slot]}</h3>
      <div class="types">${types.map((t) => `<label><input type="radio" name="type" value="${t}" ${(item?.type ?? 'empty') === t ? 'checked' : ''}> ${t === 'empty' ? 'فارغة' : TYPES[t]}</label>`).join('')}</div>
      <label>النص في الخانة<textarea name="text" rows="2">${esc(item?.text ?? '')}</textarea></label>
      <label>عدد الفترات <select name="span">${[1, 2, 3, 4].filter((n) => slot + n <= 10).map((n) => `<option ${n === (item ? item.slots.filter((x) => x >= slot).length : 2) ? 'selected' : ''}>${n}</option>`).join('')}</select></label>
      <div class="row"><button value="ok" class="primary">حفظ</button><button value="cancel">إلغاء</button></div>
    </form>`;
  const form = dialog.querySelector('form')!;
  const text = form.elements.namedItem('text') as HTMLTextAreaElement;
  form.addEventListener('change', (e) => {
    const t = (e.target as HTMLInputElement);
    if (t.name === 'type' && DUTY_TEXT[t.value as ItemType]) text.value = DUTY_TEXT[t.value as ItemType]!;
  });
  dialog.onclose = () => {
    if (dialog.returnValue !== 'ok') return;
    const type = (form.elements.namedItem('type') as RadioNodeList).value as ItemType | 'empty';
    const span = Number((form.elements.namedItem('span') as HTMLSelectElement).value);
    freezeDuties(sheet);
    const edits = ((project.decisions.edits ??= {})[sheet.staff.id] ??= {});
    // Clearing a block clears all of it; setting one covers the chosen number of slots.
    const slots = type === 'empty' && item ? item.slots : Array.from({ length: span }, (_, i) => slot + i);
    for (const s of slots) {
      edits[`${day}:${s}`] = type === 'empty' ? null : { type, text: text.value.trim() || DUTY_TEXT[type] || TYPES[type] };
    }
    commit('تم تعديل الخانة');
  };
  dialog.showModal();
}

function fixIssue(id: string, option: string): void {
  const issue = result.issues.find((i) => i.id === id);
  if (!issue) return;
  const d = project.decisions;
  if (issue.kind === 'no-underline' && issue.courseKey) (d.starts ??= {})[issue.courseKey] = option;
  else if ((issue.kind === 'both-free' || issue.kind === 'conflict') && issue.sessionId) (d.supervisor ??= {})[issue.sessionId] = option;
  else if (issue.kind === 'short-name') (d.names ??= {})[id.replace(/^short:/, '')] = option;
  commit('تم حفظ القرار');
}

async function downloadOne(id: string): Promise<void> {
  const sheet = staffSheet(id);
  if (!sheet) return;
  download(sheetFileName(sheet), (await writeSheet(sheet, project.settings)) as BlobPart, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

async function downloadAll(): Promise<void> {
  const zip = new JSZip();
  for (const sheet of result.sheets) zip.file(sheetFileName(sheet), await writeSheet(sheet, project.settings));
  download('الجداول الفردية.zip', await zip.generateAsync({ type: 'blob' }), 'application/zip');
}

function editCourses(change: (rows: StaffSheet['courses']) => void): void {
  const sheet = staffSheet(person) ?? result.sheets[0];
  if (!sheet) return;
  const rows = structuredClone(sheet.courses);
  change(rows);
  (project.decisions.courses ??= {})[sheet.staff.id] = rows;
  commit();
}

function newStaff(name = '', role: 'doctor' | 'ta' = 'doctor'): Staff {
  const bare = name.replace(/^[^\s]+\s/, '');
  return {
    id: `staff-${Date.now().toString(36)}`,
    name: name ? `${role === 'ta' ? 'م' : 'د'}/${bare}` : '',
    aliases: bare ? [bare] : [],
    programme: project.roster[0]?.programme ?? '',
    rank: role === 'ta' ? 'ta' : /^ا\.?\s*د/.test(name) ? 'professor' : /م\.?\s*د/.test(name) ? 'associate' : 'lecturer',
    advising: true,
    secondmentDay: null,
    qualityHours: 2,
  };
}

root.addEventListener('click', (e) => {
  const el = (e.target as HTMLElement).closest<HTMLElement>('[data-tab],[data-act],[data-cell]');
  if (!el) return;
  if (el.dataset.tab) {
    if ((el as HTMLButtonElement).disabled) return;
    tab = el.dataset.tab as Tab;
    message = '';
    return render();
  }
  if (el.dataset.cell) return openEditor(el.dataset.cell);
  const { act, id = '', i = '0' } = el.dataset;
  const index = Number(i);
  switch (act) {
    case 'remove-master':
      project.masters.splice(index, 1);
      return commit();

    case 'add-staff':
      project.roster.push(newStaff());
      return commit();
    case 'add-detected':
      project.roster.push(newStaff(el.dataset.name, el.dataset.role as 'doctor' | 'ta'));
      return commit('تمت الإضافة — راجع الدرجة والبرنامج ويوم الانتداب');
    case 'remove-staff':
      if (confirm(`حذف ${project.roster[index].name}؟`)) {
        project.roster.splice(index, 1);
        commit();
      }
      return;
    case 'process':
      tab = 'sheets';
      message = '';
      render();
      return window.scrollTo(0, 0);
    case 'person':
      person = id;
      tab = 'sheets';
      render();
      return document.querySelector('.person')?.scrollIntoView({ behavior: 'smooth' });
    case 'fix':
      return fixIssue(id, el.dataset.opt!);
    case 'dismiss':
      (project.decisions.dismissed ??= {})[id] = true;
      return commit();
    case 'undo-decisions':
      if (confirm('التراجع عن كل قراراتك في قائمة المشاكل؟')) {
        const { duties, edits, courses } = project.decisions;
        project.decisions = { duties, edits, courses };
        commit();
      }
      return;
    case 'reset-person':
      delete project.decisions.edits?.[id];
      delete project.decisions.duties?.[id];
      delete project.decisions.courses?.[id];
      return commit('تم إلغاء التعديلات');
    case 'add-course':
      return editCourses((rows) => rows.push({ key: `manual-${Date.now()}`, name: '', share: 0.5, weeks: '1 - 7' }));
    case 'remove-course':
      return editCourses((rows) => rows.splice(index, 1));
    case 'download-one':
      return void downloadOne(id);
    case 'download-all':
      return void downloadAll();
  }
});

root.addEventListener('change', async (e) => {
  const el = e.target as HTMLInputElement | HTMLSelectElement;
  const input = el as HTMLInputElement;
  if (el.dataset.act === 'upload' && input.files) return addFiles([...input.files]);
  if (el.dataset.act === 'sheet') {
    const m = project.masters[Number(el.dataset.i)];
    const data = fileData.get(m.fileName);
    if (!data) return;
    const master = await readMaster({ programme: m.programme, name: m.fileName, data, sheet: el.value });
    Object.assign(m, { sheet: master.sheet, sessions: master.sessions, warnings: master.warnings });
    return commit(`«${m.fileName}»: الورقة «${master.sheet}» — ${master.sessions.length} حصة`);
  }
  const row = el.closest<HTMLElement>('tr[data-i]');
  if (el.dataset.f && row) {
    const staff = project.roster[Number(row.dataset.i)] as unknown as Record<string, unknown>;
    const f = el.dataset.f;
    if (f === 'aliases') staff.aliases = el.value.split(/[،,]/).map((s) => s.trim()).filter(Boolean);
    else if (f === 'advising') staff.advising = input.checked;
    else if (f === 'secondmentDay') staff.secondmentDay = el.value === '' ? null : Number(el.value);
    else if (f === 'qualityHours' || f === 'officeHours' || f === 'advisingHours') staff[f] = Math.max(0, Number(el.value) || 0);
    else staff[f] = el.value;
    return commit();
  }
  if (el.dataset.c && row) {
    const c = el.dataset.c as 'name' | 'share' | 'weeks';
    return editCourses((rows) => {
      const target = rows[Number(row.dataset.i)];
      if (c === 'share') target.share = Number(el.value) || 0;
      else target[c] = el.value;
    });
  }
  if (el.dataset.s) {
    (project.settings as unknown as Record<string, string>)[el.dataset.s] = el.value;
    return commit();
  }
  if (el.dataset.sig) {
    project.settings.signatures[Number(el.dataset.sig)].name = el.value;
    return commit();
  }
});

root.addEventListener('dragover', (e) => {
  if ((e.target as HTMLElement).closest('#drop')) e.preventDefault();
});
root.addEventListener('drop', (e) => {
  if (!(e.target as HTMLElement).closest('#drop') || !e.dataTransfer) return;
  e.preventDefault();
  void addFiles([...e.dataTransfer.files]);
});

export function start(): void {
  render();
}
