# مولّد الجداول الفردية — Individual Timetable Generator

Reads the programme master timetables (Word / Excel) and generates every staff member's
teaching-load sheet (`جدول الأعباء التدريسية وملحقاتها`) as Excel, in the browser. Files never
leave the computer. The full brief is in [CLAUDE.md](CLAUDE.md).

## Use

```bash
npm install
npm run dev
```

Open the address it prints, then: upload masters → check the staff list → review → download.

## Develop

| Command | What it does |
|---|---|
| `npm test` | Unit tests against `samples/`, including a printed diff against the ten hand-made sheets |
| `npm run build` | Type-check and build the page into the repository root (`index.html`, `assets/`) |
| `node scripts/preview.ts [staff-id]` | Text summary of everyone, or one person's sheet |
| `node scripts/dump-sessions.ts [--staff id] [--names]` | Sessions read from the sample masters |
| `node scripts/export.ts` | Write every sheet from the sample masters into `out/` |

Logic lives in `src/core/` (pure modules, no DOM): `docx.ts` / `xlsx.ts` read a master into a grid,
`sessions.ts` + `cellText.ts` turn it into sessions, `rules.ts` applies the client's rules,
`writer.ts` writes the Excel sheet. The page is `src/ui/`.

## Deploy

GitHub Pages serves the repository root of `main` (Settings → Pages → Deploy from a branch → main, /root).
`npm run build` writes the built page (`index.html`, `assets/`) to the root; commit and push them to publish.
The page source is `web/index.html`.
