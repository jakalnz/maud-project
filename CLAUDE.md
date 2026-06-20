# MAud Clinical Assessment Ecosystem

## Project overview

A digital ecosystem for clinical supervision, student skill tracking, hours compliance, and reflective practice for the University of Auckland MAud (Master of Audiology) programme. Replaces manual Excel-based tracking with an integrated system.

See @docs/MAud_Project_Summary.docx for the full project summary.
See @docs/MAud_PebblePad_Workbook_Spec.docx for PebblePad workbook specification.
See @docs/UoA_Guidelines_for_Clinical_Assessment_v2.docx for complete behavioural descriptors per skill level.
See @docs/Objective_and_Progression_Mapping.xlsx for the authoritative skill-to-objective-to-clinic-type mapping.
See @docs/HANDOVER_PebblePad.md for PebblePad integration handover notes.

## Architecture

- **Frontend**: Static HTML/JS hosted on GitHub Pages (no framework, vanilla JS)
- **Backend**: Google Sheets + Google Apps Script (deployed as web app)
- **Authentication**: Google Identity Services One Tap (GIS); token verified server-side via local JWT decode (no tokeninfo HTTP call)
- **Portfolio**: PebblePad with ATLAS (external, not in this repo)
- **PDF generation**: Client-side via html2pdf.js (not yet wired up — see Outstanding work)

## Live URLs

- **Home**: https://jakalnz.github.io/maud-project/
- **Student Dashboard**: https://jakalnz.github.io/maud-project/src/dashboard/
- **Session Feedback & Assessment**: https://jakalnz.github.io/maud-project/src/form/
- **Log Supervised Hours**: https://jakalnz.github.io/maud-project/src/student-form/
- **Apps Script**: `https://script.google.com/macros/s/AKfycbwsc7vAdP7FcmeWK_5TPXJOxXnUBtFVmv_m1cTMp9V7T-PjhnPkbJfh-Pn1Q1-hMIHTTg/exec`
- **OAuth Client ID**: `27923847271-b8cu115ptvp7ft3dnrtl0p1mc5c5pe81.apps.googleusercontent.com`

## Folder structure

```
/
├── CLAUDE.md
├── index.html                   # Home page (navigation hub)
├── docs/                        # Specifications and reference documents
│   ├── MAud_Project_Summary.docx
│   ├── MAud_PebblePad_Workbook_Spec.docx
│   ├── SHEET_SETUP.md           # Google Sheet column specs (authoritative)
│   ├── HANDOVER_PebblePad.md    # PebblePad integration handover
│   ├── UoA_Guidelines_for_Clinical_Assessment_v2.docx
│   └── Objective_and_Progression_Mapping.xlsx
├── src/
│   ├── dashboard/
│   │   └── index.html           # Skills dashboard (GitHub Pages)
│   ├── form/
│   │   └── index.html           # Supervisor feedback form
│   ├── student-form/
│   │   └── index.html           # Student hours log
│   └── apps-script/
│       ├── Code.gs              # Main API (config, students, ratings, submit)
│       └── PlacementImporter.gs # Placement Excel hours importer
├── prototypes/                  # Reference prototypes from design phase
└── assets/
```

## What is built and working

### Home page (`index.html`)
Three cards linking to the three tools.

### Student Dashboard (`src/dashboard/index.html`)
- Google One Tap sign-in gate
- POST-based role check (`{action:'role', token}`) to Apps Script
- Supervisor with a student record sees a "View as: Supervisor / Student" toggle post-auth
- Data display (sparklines, hours gauges, skills tables) is **placeholder only** — needs wiring to API responses

### Session Feedback & Assessment (`src/form/index.html`)
- Full supervisor form: student select, date, supervisors, location
- Activities: three-column checkbox grid (Adult / Paediatric / Other)
- Sub-type selection via **bottom-sheet modal** (touch-friendly, 52px tap targets); selections stored in `_subTypes` object (not DOM checkboxes)
- Hint text under each active clinic type showing selected sub-types as comma list
- Hours: per active clinic type, observed/tested split; ORL/SLT/Simulation tagged separately
- Skills toggle (off by default): clinic-type tabs, per-skill ratings, priority/strength flags, expandable comment boxes
- Milestone: auto-detected from session date + selected student's cohort; chip display with manual override
- Submit handler exists but **PDF generation and email/Drive saving are stubs**

### Log Supervised Hours (`src/student-form/index.html`)
- Student (or supervisor with student record) self-logs hours
- Same bottom-sheet modal sub-type pattern as supervisor form
- Milestone auto-detected from session date + student's cohort
- Submit works (writes to Sessions tab via Apps Script)

### Apps Script (`src/apps-script/Code.gs`)
- `verifyToken`: local JWT decode (no external HTTP call)
- `getRole`: checks Students tab first, then Supervisors; returns `studentId` even for supervisors who are also students
- `doGet`/`doPost`: handles `config`, `students`, `ratings`, `cohort_overview`, `role` (POST only), `submitSession`, `submitStudentHours`
- `getConfig`: reads Config tab cols A–F (Year, Label, Active, S1 End, S2 End, Y2 End); returns cohorts with `s1End`/`s2End`/`y2End` date strings
- Auth on submit: checks `auth.studentId` (not `auth.role`) so supervisors with student records can submit student hours

## Google Sheet structure

See `docs/SHEET_SETUP.md` for full column specs. Tabs in order:

1. **Config** — `Year | Label | Active | S1 End | S2 End | Y2 End` ← dates must be added to live sheet
2. **Students** — `StudentID | Cohort | Name | Email`
3. **Supervisors** — `Name | Email` (just these two columns)
4. **Skills** — `SkillID | SkillName | Objective | Scope | ExpS1 | ExpS2 | ExpY2`
5. **Sessions** — one row per submitted session
6. **Ratings** — one row per skill rating per session

## Outstanding work (next session)

1. **Config tab milestone dates**: Add cols D/E/F (S1 End, S2 End, Y2 End as dates) to live Google Sheet Config tab, then **redeploy Code.gs** as a new version.

2. **PDF generation** (`src/form/index.html`): Wire up html2pdf.js in the submit handler. PDF content: student name, date, supervisor, hours table, flagged skills with comments only. No numeric ratings.

3. **Email notification** (`Code.gs`): `submitSession` writes to sheets but doesn't email the student. Add `MailApp.sendEmail()` call.

4. **Drive folder** (`Code.gs`): Create per-student Drive folder with restricted sharing; save PDF there on submission.

5. **Dashboard data display** (`src/dashboard/index.html`): Connect `ratings` and `cohort_overview` API responses to the UI — sparklines, hours gauges, skills tables are all placeholder.

6. **PebblePad integration**: See `docs/HANDOVER_PebblePad.md`. Start with PDF attachment (Option A); iframe embed and API push are parked pending UoA IT actions.

## Auth implementation notes

- GIS One Tap (`auto_select: true`) — button popup was blocked by COOP on GitHub Pages; One Tap works
- Token passed via **POST body** (not GET URL) — 1234-char token caused issues in URL params
- `verifyToken` does local JWT decode: split on `.`, base64-decode middle segment, check `aud`, `iss`, `exp`
- Role check returns `studentId` for supervisors who also appear in the Students tab — this enables the "View as" toggle and allows supervisors to submit student hours

## Milestone auto-detection

Config tab provides `s1End`, `s2End`, `y2End` per cohort. Detection logic ("up to end date"):

```js
// dateStr = YYYY-MM-DD, cohortYear = string e.g. '2026'
// Returns {code, label} or null
if (d <= new Date(cohort.s1End + 'T23:59:59')) return {code:'S1', label:'Semester 1'};
if (d <= new Date(cohort.s2End + 'T23:59:59')) return {code:'S2', label:'Semester 2'};
if (d <= new Date(cohort.y2End + 'T23:59:59')) return {code:'Y2', label:'End of Year 2'};
```

Supervisor form uses `_studentCohortMap` (built when students load) to resolve the selected student's cohort. Student form uses `_student.cohort`.

## Key design decisions

- Skill ratings are separated by clinic type — the same general skill is rated independently in each clinic context
- Numeric ratings are NOT in the student PDF — only flagged skills (priority/strength) with supervisor comments
- Sparklines show the last 5 individual session measurements, not term aggregates
- Per-student Google Drive folders use restricted sharing
- Skills toggle is off by default (quick sessions don't need full ratings)
- `_subTypes` object (not DOM checkboxes) is the source of truth for sub-type selections

## Progression milestones

THREE milestones only (S1/S2/Y2). The old T1/T2/T3/T4 system is dead — do not use it.

| Code | Name | Description |
|------|------|-------------|
| S1 | Semester 1 | End of first semester, Year 1 |
| S2 | Semester 2 | End of second semester, Year 1 |
| Y2 | End of Year 2 | End-of-course goal |

## Rating scale

1 = Absent, 2 = Emerging, 3 = Present, 4 = Developed, 5 = Consistent

- **Absent**: Unsafe or unable to do the task without specific instruction
- **Emerging**: Can attempt task, requires significant hand-holding, significant issues around timing
- **Present**: Can perform task for routine cases, minor concerns, generally minimal supervisor input, may lack confidence
- **Developed**: Appears confident, manages time effectively, needs assistance in complex cases only
- **Consistent**: Manages complex clients well, high level of competence

Full behavioural descriptors per skill are in `UoA_Guidelines_for_Clinical_Assessment_v2.docx`. These should be shown as hints in the feedback form when a rating is selected (not yet implemented).

## Clinic types

Exact strings — must match in all code, sheet data, and API responses:
- `Adult Diagnostic`
- `Paediatric Diagnostic`
- `Adult Rehabilitation`
- `Paediatric Rehabilitation`
- `Other` (observations, simulations, report writing, seminars)

## Skill applicability scopes

| Scope | Clinic types |
|-------|-------------|
| `all` | All five (including Other) |
| `all-nonobs` | Adult Dx, Paed Dx, Adult Rehab, Paed Rehab |
| `all-dx` | Adult Dx, Paed Dx |
| `adult-dx` | Adult Diagnostic only |
| `paed-dx` | Paediatric Diagnostic only |
| `adult-rehab` | Adult Rehabilitation only |
| `paed-rehab` | Paediatric Rehabilitation only |

## Complete skill mapping

### General skills

| Objective | Skill | S1 | S2 | Y2 | Scope |
|-----------|-------|----|----|----|----|
| Obj I A: Professional and Ethical Behaviour | Appearance | 4 | 4 | 4 | all |
| Obj I A: Professional and Ethical Behaviour | Clinic Maintenance | 2 | 4 | 5 | all-nonobs |
| Obj I A: Professional and Ethical Behaviour | Record Keeping | 2 | 3 | 4 | all-nonobs |
| Obj I A: Professional and Ethical Behaviour | Clinical Safety | 3 | 4 | 5 | all |
| Obj I A: Professional and Ethical Behaviour | Confidentiality | 3 | 4 | 5 | all |
| Obj I B: Interpersonal Behaviour | Interpersonal/Rapport | 2 | 3 | 4 | all |
| Obj I B: Interpersonal Behaviour | Voice and Language | 2 | 3 | 4 | all |
| Obj I C: Refining Behaviour | Refining Behaviour | 3 | 4 | 4 | all |
| Obj II: Session Preparation | Session Preparation | 2 | 4 | 4 | all |
| Obj III: Session Management | Speed and Time | 2 | 3 | 4 | all-nonobs |
| Obj IV: Reasoning and Integration | Interpretation | 1 | 2 | 4 | all-nonobs |
| Obj IV: Reasoning and Integration | Applies X-check | 2 | 3 | 4 | all-dx |
| Obj V D: Clinical Writing | Progress Notes | 1 | 2 | 3 | all-nonobs |
| Obj V D: Clinical Writing | Report Writing | 1 | 2 | 3 | all-nonobs |
| Obj V D: Clinical Writing | Audiogram | 3 | 4 | 5 | all-dx |
| Obj V C: Presentation of Findings | Explanation | 2 | 4 | 5 | all-nonobs |
| Obj V C: Presentation of Findings | Recommendations | 2 | 3 | 4 | all-nonobs |

### Adult Diagnostic skills

| Objective | Skill | S1 | S2 | Y2 |
|-----------|-------|----|----|-----|
| Obj V A: Clinical Interview | History Taking | 2 | 4 | 5 |
| Obj V B: Testing | Otoscopy | 2 | 3 | 4 |
| Obj V B: Testing | Tympanometry | 2 | 3 | 5 |
| Obj V B: Testing | ARTs | 2 | 3 | 4 |
| Obj V B: Testing | PTA (AC/BC) | 2 | 3 | 4 |
| Obj V B: Testing | PT Masking | 3 | 5 | 5 |
| Obj V B: Testing | Speech Audiometry | 3 | 4 | 4 |
| Obj V B: Testing | Speech Masking | 2 | 3 | 4 |

### Paediatric Diagnostic skills

| Objective | Skill | S1 | S2 | Y2 |
|-----------|-------|----|----|-----|
| Obj V A: Clinical Interview | History Taking | 1 | 2 | 2 |
| Obj V B: Testing | Otoscopy | 1 | 2 | 3 |
| Obj V B: Testing | VRA (Distractor) | 1 | 2 | 3 |
| Obj V B: Testing | VRA (Tester) | 1 | 2 | 3 |
| Obj V B: Testing | PLAY | 1 | 2 | 3 |
| Obj V B: Testing | KTT | 1 | 2 | 3 |
| Obj V B: Testing | OAEs | 1 | 3 | 4 |

### Adult Rehabilitation skills

| Objective | Skill | S1 | S2 | Y2 |
|-----------|-------|----|----|-----|
| Obj V A: Clinical Interview | HA Discussion | 1 | 2 | 3 |
| Obj V A: Clinical Interview | Needs Assessment (COSI/HHIA) | 1 | 2 | 3 |
| Obj IV: Reasoning and Integration | Funding | 1 | 2 | 3 |
| Obj IV: Reasoning and Integration | HA Selection | 1 | 2 | 3 |
| Obj IV: Reasoning and Integration | Troubleshooting | 1 | 2 | 3 |
| Obj V B: Testing | Impressions | 1 | 2 | 3 |
| Obj V B: Testing | Noah/HA Setup | 1 | 2 | 3 |
| Obj V B: Testing | HA Programming | 1 | 2 | 3 |
| Obj V B: Testing | Verification | 1 | 2 | 3 |
| Obj V B: Testing | Validation | 1 | 2 | 3 |
| Obj V C: Presentation of Findings | Counselling | 1 | 2 | 3 |

### Paediatric Rehabilitation skills

| Objective | Skill | S1 | S2 | Y2 |
|-----------|-------|----|----|-----|
| Obj V A: Clinical Interview | Needs Assessment | 1 | 1 | 2 |
| Obj IV: Reasoning and Integration | HA Selection | 1 | 1 | 2 |
| Obj V B: Testing | HA Programming | 1 | 1 | 2 |
| Obj V B: Testing | Verification | 1 | 1 | 3 |
| Obj V B: Testing | Validation | 1 | 1 | 2 |
| Obj V C: Presentation of Findings | Counselling | 1 | 1 | 2 |

## Objectives structure

- **Objective I**: Develop behaviours that support professional excellence
  - A: Ethical and Professional Behaviour
  - B: Interpersonal Behaviour
  - C: Refining Behaviour
- **Objective II**: Demonstrate effective session preparation
- **Objective III**: Demonstrate effective session management
- **Objective IV**: Demonstrate effective session reasoning and integration
- **Objective V**: Demonstrate effective session procedure
  - A: Clinical Interview
  - B: Testing
  - C: Presentation of Findings
  - D: Clinical Writing Skills

## Hours compliance targets

| Metric | Target | Direction |
|--------|--------|-----------|
| Total hours | >250 | Minimum |
| Simulation hours | <40 | Maximum |
| Testing hours (total) | >200 | Minimum |
| Adult diagnostic | >40 | Minimum |
| Adult rehab | >10 | Minimum |
| Paediatric diagnostic | >40 | Minimum |
| Paediatric rehab | >10 | Minimum |
| Rehab total (adult + paediatric) | >80 | Minimum |
| ORL hours | Tracked | Separate count |
| SLT hours | Tracked | Separate count |
| Clinical supervision | Tracked | Separate count |

## Apps Script API endpoints

All GET endpoints accept an optional `token` query param. Role check must use POST.

- `GET ?action=config` — returns `{ cohorts, skills }`; cohorts include `s1End`/`s2End`/`y2End`
- `GET ?action=students&cohort=2025` — returns student list for cohort
- `GET ?action=ratings&student=STU001&cohort=2025` — returns all ratings for student
- `GET ?action=cohort_overview&cohort=2025` — returns summary for all students in cohort
- `POST {action:'role', token}` — returns `{ role, email, name, studentId, cohort }`
- `POST {action:'submitSession', token, ...}` — supervisor form submission
- `POST {action:'submitStudentHours', token, ...}` — student hours form submission

## Coding conventions

- Vanilla JavaScript (no React/Vue/framework)
- CSS variables for theming (light/dark mode support)
- Mobile-responsive design
- No external dependencies except Google APIs and PDF library
- All student data flows through the Apps Script API — never hardcoded
- Clinic type names must match exactly (see above)

## Parked items (do not implement without discussion)

- Calibration exercise (blind self-rating comparison on the dashboard)
- Extended term structure beyond S1/S2/Y2
- Automated placement competency mapping (currently handled manually)
- PebblePad API push (requires UoA IT to provide OAuth2 client credentials)
- Dashboard embedding in PebblePad iframe (awaiting UoA IT whitelist request for `jakalnz.github.io`)
