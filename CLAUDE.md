# Individual Timetable Generator — Project Brief

> Put this file in the repo root as `CLAUDE.md` together with the `samples/` folder.
> Everything in here was worked out from the real files in `samples/` and from voice notes by the client (Dr. Asmaa Radi, timetable committee, Electrical Engineering Dept., Faculty of Engineering, Damanhour University). When something below conflicts with the sample files, **the sample files win** — check them.

---

## 1. The problem in one paragraph

Each engineering programme publishes a **master timetable** (جدول رئيسي) — one big grid of days × time slots × levels/sections, listing every lecture, section (تمرين) and lab (ت عملى) with the staff names. From it, the timetable committee hand-builds an **individual timetable / teaching-load sheet** (جدول فردي — "جدول الأعباء التدريسية وملحقاتها") for every staff member in the Electrical Engineering department. A staff member can teach in several programmes, so their sheet pulls from several masters. Whenever a master changes, every individual sheet goes stale and must be redone by hand. **Build a tool that reads the masters and generates every individual sheet automatically, in the exact Excel layout of the samples.**

## 2. What to build

- A **static single-page web app**, Arabic, **RTL**, deployed on **GitHub Pages** (the user already deploys this way). No backend. All processing happens in the browser; files never leave the user's computer.
- Flow: **Upload masters → Staff roster → Review & resolve → Download** (one `.xlsx` per staff member, all in a `.zip`, plus single-person download).
- Suggested stack: Vite + vanilla TypeScript (or plain JS), **JSZip** (read `.docx`, write `.zip`), **ExcelJS** (read `.xlsx` incl. rich-text underline; write styled `.xlsx`). Add a GitHub Actions workflow that builds and deploys to Pages.
- Keep parsing/rules logic in pure modules with **Node unit tests** that run against `samples/` (see §10).

## 3. Input files

| File (in `samples/masters/`) | Programme | Format | Notes |
|---|---|---|---|
| `comms_2026-27.docx` | هندسة الاتصالات والالكترونيات | Word table | Levels 1–4 × sections 1–2. **Already corrected** (missing underlines added, typos fixed). |
| `power_2026-27.docx` | هندسة القوى الكهربية | Word table | Levels 1–4 × sections 1–2. **Already corrected.** |
| `civil_2026-27.xlsx` | الهندسة المدنية | Excel | **Use only sheet `مج_1 (2)`** (2026-2027). Sheet `مج_1` is an old 2024-2025 copy — ignore it (let the user pick the sheet; default to the one whose title contains 2026). Up to 3 sections per level. Time labels have typos (`09.40`). |
| *(coming later)* preparatory (level 0) | المستوى الصفري / إعدادي | will be Word | Only a PDF exists now: `samples/reference-pdf/preparatory_level0_2026-27.pdf`. 10 sections in 2 groups. |
| *(coming later)* mechanical power | هندسة القوى الميكانيكية | will be Word | Only a PDF exists now: `samples/reference-pdf/mechanical_power_2026-27.pdf`. |

**PDF masters are not supported** — show a clear message asking for the Word/Excel original. The parser must be **layout-agnostic**: every programme has a different number of levels/sections and its own header rows, so discover structure from headers, never from hard-coded column numbers.

### 3.1 Parsing Word (`.docx`) tables
- Unzip, read `word/document.xml`. The timetable is a `w:tbl`. Use `w:tblGrid` for column count.
- Build a full occupancy grid: each `w:tc` has `w:gridSpan` (colspan) and `w:vMerge` (`restart` = origin, no value = continuation → rowspan). Honour `w:trPr/w:gridBefore` / `w:gridAfter` if present.
- Rows are irregular: some rows have no time label (sub-rows, e.g. an extra thin row) — attach them to the previous labelled slot. One label in comms reads `5:40 – 4:50` in the original (now fixed) — normalise labels by start time.
- **Header rows**: row with `السكشن` holds section numbers per column; row with `المستوي` holds `مستوى 1…4` (spanning). Empty narrow columns between levels are separators.
- **Column 0** = day (vertically merged; `السبت، الاحد، الاثنين، الثلاثاء، الاربعاء، الخميس`), **column 1** = time label.
- An origin cell → one *session*: `{programme, day, slots[], level, sections[], text, runs}`. A cell spanning several section columns = shared by those sections (usually a lecture for the whole level).
- Keep the **text runs with underline flag** (`w:rPr/w:u` with `w:val` ≠ `none`) — needed for §5.2.

### 3.2 Parsing Excel (`.xlsx`) masters
- Same model: use merged ranges for spans, find the day column and time column by their values, find section/level header rows. Read rich text with ExcelJS so per-run `font.underline` is available.

### 3.3 Time slots (fixed, 10 per day)
`09:00-09:50, 09:50-10:40, 10:45-11:35, 11:35-12:25, 12:30-01:20, 01:20-02:10, 02:15-03:05, 03:05-03:55, 04:00-04:50, 04:50-05:40`. Days: Saturday → Thursday. **Each slot = 1 hour** for load calculations. Map any label to a slot by its start time (tolerate `09.40`, `08:50` etc.; anything that cannot be mapped → warning).

## 4. Reading a session's text

Typical cells (Arabic, spacing/diacritics/tatweel are inconsistent):
```
محــاضـــرة مبـــادئ بــث المعلـومــــات (ELC 151) (د. أحمد سالم & د. <u>اسماء راضى</u>) / مدرج 7 (No. 722)
ت عملى الاتصالات التناظرية (م. جوستينا) (ا.م.د. هند السيد& د اسماء راضى) /  Lab 729
تمرين رياضيات 3 (م.كريم مصطفى) (د.أسماء عبدالرحيم& د اسماء راضى)  قاعة 702
مقرر اختيارى متطلب جامعة (HUM E)(القيادة والادارة) / (د احمد سالم) (د احمد السيد ) / مدرج مدنى (No. 327)
```
- **Normalise** before matching: remove tatweel `ـ`, unify `أإآ→ا`, `ى→ي`, `ة→ه`, collapse spaces.
- **Type**: contains `محاضرة` or is an elective `مقرر اختيارى` → **lecture**. Contains `ت عملى` / `عملى` / `معمل` / `تمرين` / `م برمجة` / `ورش` → **section/lab**. `أنشطة طلابية` → ignore. `مشروع التخرج` → flag for review.
- **Course**: code regex like `[A-Z]{3}\s*[0-9A-Z]{3,4}` (ELC 151, ELP3E1, MAT001, HUM E). Sections usually omit the code → link a section to its lecture by normalised course title within the same programme (and level). Keep the title as written for the output.
- **Room**: `مدرج …`, `قاعة …`, `Lab …`, `معمل …`, `(No. …)` — keep the tail after the names as the room text.
- **Staff names**: title prefixes `د.` `د/` `د` `ا.د.` `أ.د` `ا.م.د.` `أ. م. د.` (doctors) and `م.` `م/` `م.م` (teaching assistants). Separators `&`, `)(`. Do **not** rely on free-form name extraction alone — **match against the roster aliases** (§6), and list any unmatched `د…` tokens on the review screen so the user can map them.
- **Short names** (`د. اسماء`, `د. ايمان`, `د. محمد`) appear in preparatory/civil labs. Resolve them from the doctors of the same course's lecture; otherwise use the roster alias table; otherwise ask on the review screen. Known: in the preparatory timetable `د. اسماء` = **د. أسماء عبدالرحيم**.

## 5. Rules (confirmed by the client)

### 5.1 Lectures
- The doctor gets the lecture in their sheet at the same day/slots with the room.
- **A lecture counts its full slot hours** (2 slots → 2 hours) even if the doctor teaches only half the semester.

### 5.2 Shares and weeks (the underline rule)
- Lecture with **two doctors**: the **underlined** doctor teaches **weeks 1–7**, the other **weeks 9–15**; each gets share **0.5** (`نسبة المشاركة`).
- Lecture with **one doctor**: share **1**, weeks **1–15**.
- Underline detection: a doctor is "underlined" if most of the letters of their name are in underlined runs (the original files had stray underlines on single letters/spaces).
- Two doctors and **no underline** → review-screen item ("who starts?").
- **Only lecture cells decide weeks.** Ignore underlines on section/lab cells.
- The user can override share/weeks per (doctor, course) on the review screen.

### 5.3 Sections and labs (supervision — اشراف)
- Every section/lab cell lists **both** course doctors, but in reality **only one supervises each session**. Rule from the client: **each doctor takes the sessions that fall when they have no lecture; the other doctor takes the rest.**
- Algorithm: for each session, candidates = listed doctors who are in the department roster. Remove any candidate who is busy then (lecture, secondment day, or another assigned session). One left → assign. Both free → **heuristic**: keep the same doctor on the same section number of that course across its sessions (in the samples Dr. Asmaa Radi took section 1 of Analog Communications, Dr. Hend section 2), then balance toward each doctor's load target. None free → conflict on the review screen.
- Level-1 courses are shared across programmes (ELC 151 and MAT 101 appear in both comms and power), each with its **own** sections — they are separate sessions.
- In the sheet a supervised session is written `اشراف ` + session text (e.g. `اشراف ت عملى الاتصالات التناظرية / Lab 729`).
- **Teaching assistant** (`م.` name in the cell): the TA gets **every** session they are named in; share 1, weeks 1–15.

### 5.4 Non-timetable duties (entered per person, remembered)
| Duty | Arabic | Counted in load? | Notes |
|---|---|---|---|
| Quality | جودة | **No** (shown and in daily total, but excluded from النصاب) | Everyone. Default 2 hours. |
| Office hours | ساعات مكتبية | Yes | Usually 4 or 6. |
| Academic advising | ارشاد أكاديمي | Yes | Only for advisors (roster flag). Usually 4. |
| Secondment day | انتداب | — | Whole day blocked, written as `انتداب` across the row. |

**Minimum counted load (lectures + supervision + office hours + advising):**

| Rank | Arabic | Minimum |
|---|---|---|
| Professor | أستاذ دكتور (ا.د) | 25 |
| Associate professor | أستاذ مساعد (ا.م.د) | 27 |
| Lecturer | مدرس (د) | 29 |
| Teaching assistant | مدرس مساعد / معيد (م) | 33 |

If teaching does not reach the minimum, **suggest** office-hour and advising blocks (advisors only) in free slots (prefer 2-slot blocks, never on the secondment day, avoid Thursday unless needed) until it does. The user can move/delete any block in an editable grid before export.

## 6. Staff roster (seed data — editable in the UI, saved)

Only people **in the Electrical Engineering department** get a sheet (plus TA Tamer). People from other departments (e.g. د. مايسة, د. محمد خيرت, د. أحمد نبيوه) are ignored even if they share a course.

| Name on sheet | Programme (القسم line) | Rank (الدرجة) | Advising | Secondment |
|---|---|---|---|---|
| د/أسماء راضي | هندسة الاتصالات والالكترونيات | مدرس | yes | Saturday |
| د/أحمد ابراهيم سالم | هندسة الاتصالات والالكترونيات | مدرس | **no** | — |
| د/مي حلمي | هندسة الاتصالات والالكترونيات | مدرس | yes | Wednesday (per her sheet) |
| د/هند علي السيد | هندسة الاتصالات والالكترونيات | أستاذ مساعد | yes | Sunday (per her sheet) |
| د/أحمد السيد متولي | هندسة القوي الكهربية | مدرس | yes | Wednesday |
| د/ أسماء عبدالرحيم إبراهيم السقعان | هندسة القوى الكهربية | مدرس | yes | Wednesday |
| ا.د/محمود عبدآمين السد | هندسة القوى الكهربية | استاذ دكتور | **no** | Monday (per his sheet) |
| ا.م.د/جمعه فهمى عبدالنبى | هندسة القوى الكهربية | استاذ مساعد | **no** | — |
| د/ايمان احمد عوض | هندسة القوى الكهربية | مدرس | yes | Thursday |
| م/تامر الشرقاوى | هندسة القوى الكهربية | مدرس مساعد | yes (per his sheet) | — |
| د. محمد سعيد | ? | ? | yes | ? — **confirm membership** |
| د. ايمان شوقى | ? | ? | probably yes | ? — **confirm** |

Other names found in the comms/power masters that may or may not be department staff — show them as "detected, not in roster" so the user can add them: ا.م.د احمد المليجى، د. بسنت سمير، ا.د. عبد الوهاب العيسوى، د. سيد محمد احمد، د حمدى الشامى، د احمد بكير، ا.م.د جمعة عثمان (head of dept).

Each roster entry needs: display name, **aliases** (spelling variants — e.g. `اسماء راضى`, `أسماء راضي`; `هند السيد`, `.م.د هند السيد`; `جمعه فهمى`, `جمعة فهمى`; `أحمــد السيد`, `احمد السيد`, `اخمد السيد`), programme line, rank, advising flag, secondment day, quality hours, office/advising targets. **Beware similar names**: أسماء راضي vs أسماء عبدالرحيم; أحمد سالم vs أحمد السيد; ايمان عوض vs ايمان شوقى; جمعة فهمى vs جمعة عثمان.

## 7. Output: the individual sheet (`.xlsx`)

Reproduce `samples/expected/asmaa_radi.xlsx` (doctor) and `samples/expected/tamer_elsharkawy_TA.xlsx` (TA). **Open them and copy the layout, column widths, row heights, fonts (Calibri, bold, large sizes), borders, merges and page setup** (RTL sheet, landscape, A4 paper size 9, fit to page). Structure (doctor sheet with 6 courses):

| Rows | Content |
|---|---|
| D3:K4 | `جدول الأعباء التدريسية وملحقاتها` |
| D5:K5 | `الفصل الدراسي الأول 2027-2026` (semester/year configurable) |
| B7:G8 | `الإسم: <name>` |
| J8:M9 | `أسماء المقررات المشارك فيها` |
| row 10 | B `القسم:` · C10:E10 programme · F `الدرجة:` · G rank · J10:K10 `المقرر` · L `نسبة المشاركة` · M `الاسابيع` |
| rows 11… | one row per course: J:K course name (with code), L share (0.5 / 1), M weeks (`1 - 7`, `9 - 15`, `1 - 15`). **Number of rows varies** (6 for Asmaa Radi, 9 for Ahmed El-Sayed) — everything below shifts. |
| next row + 1 | time header C…L (10 slots), M `المجموع` |
| 6 rows | السبت … الخميس: B day, C…L slots, M daily total. A 2-slot item is one merged cell over 2 columns. Secondment = `انتداب` merged across C:L. |
| summary header | doctor: `المحاضرات` / `الساعات المكتبية` / `ساعات الإشراف` / `ساعات الإرشاد الأكاديمي` / `الجــــودة` / `المجموع` (pairs merged C:D, E:F, G:H, I:J, K:L). **TA: first column is `تمرين/عملي`** (section hours) and supervision is 0. |
| values row | the numbers; M = `SUM` of all incl. quality |
| `النصاب` row | the counted values (quality excluded); M = `SUM` excluding quality. *(Some samples have stale numbers here — see open questions.)* |
| signatures | `لجنة الجدول` / `رئيس مجلس القسم` / `وكيل الكلية لشئون التعليم والطلاب` / `عميد الكلية` and the names under them — **configurable** (the head-of-dept name differs between the comms and power samples: `أ. م. د. جمعة عثمان` vs `أ. م. د. جمعة فهمى`). |

- **Daily total** = number of filled slots that day (all types, incl. quality).
- Cell text: lecture → `محاضرة <course> (<code>) <room>`; supervision → `اشراف <session text>`; duties → `جودة` / `ساعات مكتبية` / `ارشاد أكاديمي`.
- Use one consistent fill per type (the samples are inconsistent): lecture `92D050`, supervision `FFFF00`, office/advising/quality `99FFCC`, secondment none.
- File name: `جدول فردي - <name>.xlsx`; zip: `الجداول الفردية.zip`.

## 8. Review screen (must have)
- List of problems to resolve, each with a one-click fix: lectures with 2 doctors and no underline; section sessions where both doctors are free (show the proposal, allow swap); conflicts (doctor in two places); unmatched names / short names; staff below their minimum load; unparseable cells or time labels.
- Per-person preview of the generated grid (editable: click a cell → set lecture/supervision/office/advising/quality/empty).
- All decisions saved (`localStorage`) **and** exportable/importable as a JSON "project file" so she can move between computers. When masters are re-uploaded, keep prior decisions that still apply and only ask about what changed.

## 9. Known data quirks (handle, don't crash)
- Names written many ways (`د.`, `د/`, `د `, `ا.م.د.`, `.م.د`, tatweel, `أحمــد`); `& &`; missing second name (`(د. أحمد السيد&)`).
- Same course code reused for different courses across programmes (e.g. `ELP 331` = الآلات المتزامنة in power, الات القوى الكهربية in mechanics; her sheet also lists دوائر كهربية والكترونية as ELP 331 while the mechanics PDF says ELP 211) → identify courses by **programme + code + title**, not code alone.
- In the mechanics timetable the course `دوائر كهربية والكترونية` lists only د/ مي حلمي; the client will add her own name (Asmaa Radi 1–7, Mai Helmy 9–15).
- Labels like `Lab 714 & قاعة 702` contain `&` — don't treat that as a name separator.

## 10. Testing against the samples
`samples/expected/` holds 10 hand-made sheets. Write a test that generates each person from the masters available and diffs it against the sample: courses+share+weeks, every grid cell (type + course), daily totals, summary values. **The samples were made by hand from a slightly older master, and some people teach in programmes whose masters are still PDF-only (preparatory, mechanics)**, so a 100% match is not expected — print a readable diff per person and aim for: all lectures and weeks match; supervision matches where the source master is available. Reference values from the samples:

| Person | Lectures | Office | Supervision | Advising | Quality | Notes |
|---|---|---|---|---|---|---|
| Asmaa Radi | 12 | 4 | 15 | 4 | 2 | daily 10/9/8/10; uses comms + power + mechanics |
| Ahmed Salem | 10 | 4 | 21 | 4 | 2 | no advising per client, but his sheet shows 4 — confirm |
| Mai Helmy | 8 | 5 | 12 | 4 | 2 | |
| Ahmed El-Sayed | 18 | 4 | 9 | 4 | 2 | |
| Hend El-Sayed | 10 | 4 | 10 | 4 | 2 | |
| Asmaa Abdel-Rahim | 15 | 5 | 9 | 6 | 5 | also civil + preparatory |
| Mahmoud El-Sadd | 10 | 6 | 10 | 6 | 4 | also preparatory (Math 1) |
| Gomaa Fahmy | 14 | 8 | 13 | 1 | 4 | also preparatory (Math 1), mechanics |
| Eman Awad | 16 | 6 | 11 | 5 | 2 | also preparatory (Computers & Programming) |
| Tamer (TA) | 25 (sections) | 0 | 0 | 6 | 5 | |

## 11. Open questions (build so these are settings, not code changes)
1. **`النصاب` row** — **ANSWERED (client, 2026-10-08): keep it as in the individual sheets** → counted values (lectures / office / supervision / advising), quality cell empty, total = `SUM` without quality; editable per person. Background: in 5 samples it equals the counted values; in 5 others it holds different numbers (e.g. Gomaa 6/8/20/0). Default: copy counted values; ask the client what it should be.
2. **ANSWERED (corrected 2026-10-09): د. محمد سعيد and د. ايمان شوقى are from other faculties — NOT in the roster, no sheets.** Still open: membership of the "detected" names in §6.
3. Advising — **ANSWERED (client, twice; final 2026-10-09): د. أحمد سالم, ا.د. محمود السد and ا.م.د. جمعه فهمى have NO advising**, whatever their old sheets show. Everyone must reach the minimum load; the tool tops up with office hours (and advising for advisors).
4. Who starts Power Network Analysis (ELP 421): the corrected master underlines **Gomaa Fahmy** (matches both people's sheets); the client said Mahmoud in a voice note.

## 12. Suggested build order
1. Docx parser → normalised sessions (print a debug table for comms and power; check against the reference PDFs visually).
2. Name matching + roster; list detected names.
3. Rules engine (lectures, weeks, section split, TA).
4. Excel writer reproducing the template exactly; compare with `asmaa_radi.xlsx` in Excel/LibreOffice.
5. Tests vs `samples/expected/`.
6. UI (upload → roster → review → download), persistence, xlsx-master parser (civil).
7. GitHub Pages workflow + short Arabic usage guide on the page.

## 13. Decisions after the first build (client / owner, 2026-10-09)
- **`samples/expected` are the target sheets.** Each person's office / advising / quality blocks, their hours, and their lectures/labs from the PDF-only masters (mechanics, preparatory) are remembered from those sheets as roster *presets* (`src/core/seedPresets.ts`, regenerate with `node scripts/seed-from-samples.ts`). Presets are placed only where the masters leave the slot free; a master that later covers the same slot wins.
- Lab supervision when both doctors are free: leave a doctor's preset duty time alone, then split the course's sessions evenly between the two doctors, then keep the same doctor on the same section, then balance total load.
- Everyone must reach the minimum load; the tool tops up with office hours and advising.
- Commits carry the owner's name only (no co-author lines).
- Both-free order (revised 2026-10-09, supersedes the line above): 1) leave preset duty time alone, 2) prefer the doctor who already comes in that day, 3) same doctor stays with the same section, 4) split the course's sessions evenly, 5) the doctor who does not hold the other section, 6) lower total load. With this, Asmaa Radi, Mai Helmy and Ahmed Salem match their sheets' supervision fully.

## 14. Client answers by voice note (Dr. Asmaa Radi, 2026-10-09) — these supersede §13 where they differ
- **Order she works in by hand:** 1) lectures, 2) supervision of sections/labs, 3) academic advising, 4) office hours. Office hours and advising are *not* reserved first; they go into whatever is free afterwards. Advising goes on days the doctor is in the faculty; office hours may go on any other day (e.g. Wednesday or Thursday), never on the secondment day.
- **Supervision when both doctors are free:** two doctors with two labs → each takes one; share evenly; prefer the doctor who is in the faculty that day (has lectures that day). Implemented order: in the faculty that day → same doctor stays with the same section → even split → not holding the other section → lower load. Duty presets no longer block a slot for supervision; they only tell which days a person attends.
- **`النصاب` row:** the comms sheets are the correct ones. The power sheets (Mahmoud, Asmaa Abdelrahim, Eman, Gomaa) were made by Eng. Tamer and their النصاب numbers are not accurate — do not reproduce them.
- **Masters win over the old individual sheets** (rooms, the civil MAT 101 lecture — which the old sheet had missed).
- **Who starts:** طرق عددية (MAT 201) → د. أسماء عبدالرحيم; اختيارى 1 نظم القياسات (ELP 3E1) → she *believes* د. سيد محمد أحمد (not certain). Coded as `CLIENT_STARTS` in `src/core/rules.ts`; an underline in the master or a review choice overrides it.
- **Errors in the power master she will fix herself:** the orphan `/ Lab 825` cell (Sunday, level 1, section 2) and م. تامر listed in two labs at once.
- **No listed doctor free for a lab** (same voice notes): a doctor takes a lab only when free and in the faculty that day, so such a session gets no department supervisor. It is not a review item — the tool no longer lists these.
- **نظم القياسات start:** she said to confirm it with د. أسماء عبدالرحيم.
- **Clarification (voice note, 2026-10-09 11:28):** "splitting the labs" means section 1 of a course goes to one doctor and section 2 to the other — a single section is never split. If one section cannot be supervised by either (it clashes with their lectures) and only the other is open to both, it goes to **whoever still needs hours to complete their minimum load** (example: Dr. Hend, minimum 27, gets it if Dr. Asmaa Radi has already passed 29). Implemented as the last tie-break (`load`, counting the person's usual office and advising hours).
- نظم القياسات: Dr. Asmaa Abdelrahim's own sheet lists her weeks as 9–15, so Dr. Sayed starts — matches `CLIENT_STARTS`.
- Mechanics and preparatory masters: she will send the Word/Excel files as soon as she receives them.

## 15. Preparatory master received (2026-10-09)
- `samples/masters/preparatory_2026-27.docx` — converted with Word from the client's `.doc` (the site refuses `.doc` with instructions to re-save as `.docx`). The timetable sits inside a text box; one table, sections 1–10 in two groups, no level row (level = null).
- Conventions there: `ت <course>` = تمرين, `م <course>` = معمل (`م برمجة` is the lab of `حاسبات وبرمجة`), `أنشطة وندوات` is ignored, section cells use first names only (`د. محمود & د. جمعة`, `د. ايمان & د. اسماء`) — resolved from the course's lecture.
- Each lecture is given twice (once per group). MAT 001 and HUM 002 have no underline; who starts is taken from the hand-made sheets (`CLIENT_STARTS`): جمعه فهمى and أسماء عبدالرحيم.
- Remembered (preset) lectures are dropped automatically once a master provides the same course.
- Open: Dr. Gomaa's MAT 001 lecture for group 1 (Sunday 10:45) clashes with his ELP 441 lecture in power.

## 16. Corrected power master received (2026-10-09)
- `samples/masters/power_2026-27.docx` is now the client's new file: م. تامر removed from the Monday `اختيارى 4 الالات الخاصة` lab (his double booking is gone). The orphan `/ Lab 825` cell is still there.
- **Her working copy has no underlines on six lectures** (ELP 351, 441, 421, 422, 331, 321) that the earlier corrected copy had. That copy is kept as `samples/reference-docx/power_2026-27_underlined.docx`, and `src/core/seedStarts.ts` (regenerate with `node scripts/seed-starts.ts`) remembers who starts each two-doctor lecture. Order of trust: review choice → underline in the uploaded master → same course in another uploaded master → `CLIENT_STARTS` → `SEED_STARTS`.
