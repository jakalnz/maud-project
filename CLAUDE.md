# MAud Clinical Assessment Ecosystem

## Project overview

A digital ecosystem for clinical supervision, student skill tracking, hours compliance, and reflective practice for the University of Auckland MAud (Master of Audiology) programme. Replaces manual Excel-based tracking with an integrated system.

See @docs/MAud_Project_Summary.docx for the full project summary.
See @docs/MAud_PebblePad_Workbook_Spec.docx for PebblePad workbook specification.
See @docs/UoA_Guidelines_for_Clinical_Assessment_v2.docx for complete behavioural descriptors per skill level.
See @docs/Objective_and_Progression_Mapping.xlsx for the authoritative skill-to-objective-to-clinic-type mapping.

## Architecture

- **Frontend**: Static HTML/JS hosted on GitHub Pages (no framework, vanilla JS)
- **Backend**: Google Sheets + Google Apps Script (deployed as web app)
- **Authentication**: Google OAuth via Google Identity Services
- **Portfolio**: PebblePad with ATLAS (external, not in this repo)
- **PDF generation**: Client-side via jsPDF or html2pdf.js

## Folder structure

```
/
├── CLAUDE.md
├── docs/                        # Specifications and reference documents
│   ├── MAud_Project_Summary.docx
│   ├── MAud_PebblePad_Workbook_Spec.docx
│   ├── SHEET_SETUP.md
│   ├── UoA_Guidelines_for_Clinical_Assessment_v2.docx
│   └── Objective_and_Progression_Mapping.xlsx
├── src/
│   ├── dashboard/
│   │   └── index.html           # Skills dashboard (GitHub Pages)
│   ├── form/
│   │   └── index.html           # Supervisor feedback form (GitHub Pages)
│   └── apps-script/
│       ├── Code.gs              # Main API (config, students, ratings)
│       └── PlacementImporter.gs # Placement Excel hours importer
├── prototypes/                  # Reference prototypes from design phase
└── assets/
```

## Key design decisions

- Skill ratings are separated by clinic type. The same general skill (e.g. Interpersonal) is rated independently in each clinic context.
- The feedback form supports mixed clinic days: supervisor checks activities across types, hours entered per clinic type, skills rated via clinic-type tabs.
- Numeric ratings are NOT in the student PDF — only flagged skills (priority/strength) with supervisor comments. Students see ratings on the dashboard or via calibration exercises.
- Sparklines show the last 5 individual session measurements, not term/semester aggregates.
- Per-student Google Drive folders use restricted sharing ("Restricted — only people with access").
- Dashboard uses Google OAuth: student email → own data only; supervisor email → full cohort.
- The feedback form has a skills toggle (off by default for quick sessions, on for comprehensive assessment).

## Progression time points

There are THREE progression milestones (not four terms):

| Code | Name | Description |
|------|------|-------------|
| S1 | Semester 1 | End of first semester, Year 1 |
| S2 | Semester 2 | End of second semester, Year 1 |
| Y2 | End of Year 2 | End-of-course goal |

Note: The old T1/T2/T3/T4 system is superseded. All dashboard, form, and Sheet references must use S1/S2/Y2.

## Rating scale

1 = Absent, 2 = Emerging, 3 = Present, 4 = Developed, 5 = Consistent

General template:
- **Absent**: Unsafe or unable to do the task without specific instruction
- **Emerging**: Can attempt task, requires significant hand-holding, significant issues around timing
- **Present**: Can perform task for routine cases, minor concerns, generally minimal supervisor input, may lack confidence
- **Developed**: Appears confident, manages time effectively, needs assistance in complex cases only
- **Consistent**: Manages complex clients well, high level of competence

Full behavioural descriptors for each skill are in UoA_Guidelines_for_Clinical_Assessment_v2.docx. These should be shown as hints in the feedback form when a rating is selected.

## Clinic types

Five clinic types used throughout the system:
- Adult Diagnostic
- Paediatric Diagnostic
- Adult Rehabilitation
- Paediatric Rehabilitation
- Other (observations, simulations, report writing, seminars)

## Skill applicability by clinic type

Skills have different applicability scopes:

| Scope | Meaning | Clinic types |
|-------|---------|-------------|
| All (General) | Applies everywhere including observations and simulations | Adult Dx, Paed Dx, Adult Rehab, Paed Rehab, Other |
| All non-observation (Audiology) | Applies to all audiology clinics but NOT observations/simulations | Adult Dx, Paed Dx, Adult Rehab, Paed Rehab |
| All Diagnostic | Applies to diagnostic clinics only | Adult Dx, Paed Dx |
| Adult Diagnostic | Adult diagnostic only | Adult Dx |
| Paediatric Diagnostic | Paediatric diagnostic only | Paed Dx |
| Adult Rehabilitation | Adult rehab only | Adult Rehab |
| Paediatric Rehabilitation | Paediatric rehab only | Paed Rehab |

## Complete skill mapping

### General skills (rated per clinic type context)

| Objective | Skill | S1 | S2 | Y2 | Applicable to |
|-----------|-------|----|----|----| --------------|
| Obj I A: Professional and Ethical Behaviour | Appearance | 4 | 4 | 4 | All (General) |
| Obj I A: Professional and Ethical Behaviour | Clinic Maintenance | 2 | 4 | 5 | All non-observation |
| Obj I A: Professional and Ethical Behaviour | Record Keeping | 2 | 3 | 4 | All non-observation |
| Obj I A: Professional and Ethical Behaviour | Clinical Safety | 3 | 4 | 5 | All (General) |
| Obj I A: Professional and Ethical Behaviour | Confidentiality | 3 | 4 | 5 | All (General) |
| Obj I B: Interpersonal Behaviour | Interpersonal/Rapport | 2 | 3 | 4 | All (General) |
| Obj I B: Interpersonal Behaviour | Voice and Language | 2 | 3 | 4 | All (General) |
| Obj I C: Refining Behaviour | Refining Behaviour | 3 | 4 | 4 | All (General) |
| Obj II: Session Preparation | Session Preparation | 2 | 4 | 4 | All (General) |
| Obj III: Session Management | Speed and Time | 2 | 3 | 4 | All non-observation |
| Obj IV: Reasoning and Integration | Interpretation | 1 | 2 | 4 | All non-observation |
| Obj IV: Reasoning and Integration | Applies X-check | 2 | 3 | 4 | All Diagnostic |
| Obj V D: Clinical Writing | Progress Notes | 1 | 2 | 3 | All non-observation |
| Obj V D: Clinical Writing | Report Writing | 1 | 2 | 3 | All non-observation |
| Obj V D: Clinical Writing | Audiogram | 3 | 4 | 5 | All Diagnostic |
| Obj V C: Presentation of Findings | Explanation | 2 | 4 | 5 | All non-observation |
| Obj V C: Presentation of Findings | Recommendations | 2 | 3 | 4 | All non-observation |

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

- `?action=config` — returns cohorts, skill definitions, progression milestones
- `?action=students&cohort=2025` — returns student list
- `?action=ratings&student=STU001&cohort=2025` — returns all ratings
- `?action=cohort_overview&cohort=2025` — returns summary for all students

## Feedback form structure

The supervisor form has these sections in order:
1. **Session header**: Student, date, supervisor(s), location
2. **Activities**: Three-column layout (Adult / Paediatric / Other), supervisor checks what was done
3. **Hours**: Per active clinic type, observed/tested split. ORL, SLT, Simulation tagged separately.
4. **Skills toggle** (off by default): When on, shows clinic-type tabs. Each tab shows general skills + clinic-specific skills with independent ratings. Each skill has priority/strength flags with expandable comment boxes. Behavioural descriptor hints shown on rating selection.
5. **Feedback**: Done well, To improve, General comments
6. **Submit**: Writes to Google Sheet, generates PDF (without numeric ratings — only flagged skills with comments), saves PDF to student's Drive folder, emails notification.

## Coding conventions

- Vanilla JavaScript (no React/Vue/framework)
- CSS variables for theming (light/dark mode support)
- Mobile-responsive design
- No external dependencies except Google APIs and PDF library
- All student data flows through the Apps Script API — never hardcoded
- Clinic type names must match exactly: "Adult Diagnostic", "Paediatric Diagnostic", "Adult Rehabilitation", "Paediatric Rehabilitation", "Other"

## Parked items (do not implement without discussion)

- Calibration exercise (blind self-rating comparison on the dashboard)
- Extended term structure beyond S1/S2/Y2 (if programme structure changes)
- Automated placement competency mapping (currently handled manually)
- PebblePad API integration (currently manual workflow)
- Dashboard embedding in PebblePad (awaiting whitelist request)
