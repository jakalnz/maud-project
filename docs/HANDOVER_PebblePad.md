# MAud Project — PebblePad Integration Handover

**Date**: 2026-06-20  
**Repo**: https://github.com/jakalnz/maud-project  
**GitHub Pages**: https://jakalnz.github.io/maud-project/  
**Apps Script URL**: `https://script.google.com/macros/s/AKfycbwsc7vAdP7FcmeWK_5TPXJOxXnUBtFVmv_m1cTMp9V7T-PjhnPkbJfh-Pn1Q1-hMIHTTg/exec`

---

## What has been built (in this repo)

### Three web pages (GitHub Pages, vanilla JS)

| Page | URL path | Purpose |
|------|----------|---------|
| Home | `/` | Navigation hub with card links |
| Student Dashboard | `/src/dashboard/` | Student skill progression, hours summary, session history. Google OAuth gate — student sees own data only; supervisor sees full cohort. |
| Session Feedback & Assessment | `/src/form/` | Supervisor submits session details, hours, activity sub-types, and skill ratings (with priority/strength flags). |
| Log Supervised Hours | `/src/student-form/` | Student self-logs hours not entered by a supervisor. |

### Backend (Google Sheets + Apps Script)

- **Spreadsheet tabs**: Config, Students, Skills, Sessions, Ratings
- **Config tab columns**: Year | Label | Active | S1 End | S2 End | Y2 End (milestone dates per cohort)
- **API endpoints** (`?action=`): `config`, `students`, `ratings`, `cohort_overview`; POST `role` for auth
- **Auth**: Google Identity Services One Tap → JWT verified server-side; role resolved from Students/Supervisors tabs
- See `docs/SHEET_SETUP.md` for full column specs

### Key design decisions already locked in
- Skill ratings are **per clinic type** (same skill rated independently in Adult Dx vs Paed Dx etc.)
- Numeric ratings are **not** in student PDFs — only flagged skills (priority/strength) with supervisor comments
- **Progression milestones**: S1, S2, Y2 only (not T1–T4)
- Sparklines on dashboard = last 5 individual session measurements
- Five clinic types: Adult Diagnostic, Paediatric Diagnostic, Adult Rehabilitation, Paediatric Rehabilitation, Other

---

## What PebblePad needs to do

PebblePad is the **student portfolio** layer. It is external to this repo and managed via the University's ATLAS account. The spec is in `docs/MAud_PebblePad_Workbook_Spec.docx`.

### Current workflow (manual)
1. Supervisor submits form → data lands in Google Sheet → PDF generated → saved to student's Drive folder
2. Student opens PebblePad workbook → manually copies in key data or attaches PDF
3. Coordinator views ATLAS to review student portfolios at progression milestones

### Integration goal
Replace the manual copy-paste step. Options in rough order of feasibility:

| Approach | How | Notes |
|----------|-----|-------|
| **A. PDF attachment** | Form generates PDF → student downloads/attaches to PebblePad workbook | Already partially built (html2pdf.js planned). Manual but structured. |
| **B. Link embed** | Embed dashboard in PebblePad via iframe | Awaiting whitelist request (UoA IT must whitelist `jakalnz.github.io`). Parked. |
| **C. PebblePad API** | Push data from Apps Script into PebblePad workbook fields via API | PebblePad has a REST API but requires OAuth2 client credentials from UoA PebblePad admin. Not yet attempted. |

### The PebblePad workbook structure (from spec doc)
The workbook has these sections — each maps to data we already collect:
- **Student info**: name, cohort, student ID
- **Hours summary**: total, by clinic type, simulation cap, ORL/SLT
- **Skills progression**: per skill, per milestone — current rating vs expected target
- **Flagged feedback**: priority skills (with comments), strength skills (with comments)
- **Supervisor signatures**: per session

---

## Files to read before starting

```
docs/MAud_PebblePad_Workbook_Spec.docx   ← what the workbook should contain
docs/MAud_Project_Summary.docx           ← full project context and stakeholder requirements
docs/SHEET_SETUP.md                      ← Google Sheet column definitions (authoritative)
src/apps-script/Code.gs                 ← API; getConfig, getRatings, getStudents, submitSession
CLAUDE.md                                ← skill definitions, clinic types, rating scale, scope rules
```

---

## Pending technical work (non-PebblePad) to hand back

These are not yet complete in the repo and will need to be finished:

1. **PDF generation** (`src/form/index.html`): The submit handler calls a `generatePDF()` stub. Need to wire up html2pdf.js or jsPDF. PDF should include: student name, date, supervisor, hours table, flagged skills with comments. Numeric ratings excluded.

2. **Email notification** (`Code.gs`): `submitSession` writes to sheets but doesn't send email yet. Should email student when a session is submitted.

3. **Drive folder creation** (`Code.gs`): Per-student Drive folder with restricted sharing — planned but not implemented.

4. **Dashboard data visualisation** (`src/dashboard/index.html`): Auth and role switching works. The data display (sparklines, hours gauges, skills tables) is largely placeholder — needs connecting to the `ratings` and `cohort_overview` API responses.

5. **Config tab milestone dates**: The coordinator needs to add columns D/E/F (S1 End, S2 End, Y2 End as dates) to the live Config tab in the Google Sheet, then redeploy Code.gs.

---

## Architecture constraints

- **No framework** — vanilla JS only (no React/Vue)
- **No external dependencies** except Google APIs and a PDF library
- Clinic type name strings must match exactly: `"Adult Diagnostic"`, `"Paediatric Diagnostic"`, `"Adult Rehabilitation"`, `"Paediatric Rehabilitation"`, `"Other"`
- All student data flows through the Apps Script API — never hardcode student data in HTML
- OAuth Client ID: `27923847271-b8cu115ptvp7ft3dnrtl0p1mc5c5pe81.apps.googleusercontent.com`

---

## Parked items (do not implement without discussion)

- Calibration exercise (blind self-rating comparison on dashboard)
- Extended milestone structure beyond S1/S2/Y2
- Automated placement competency mapping
- PebblePad API push (Option C above) — depends on UoA IT providing client credentials
- Dashboard embed in PebblePad iframe — depends on whitelist request
