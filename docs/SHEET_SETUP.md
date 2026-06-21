# Google Sheets Setup Guide

## Sheet structure

Create a Google Sheet with **five tabs** in this order:

1. `Config` — cohort list
2. `Students` — student roster
3. `Skills` — skill definitions and progression targets
4. `Sessions` — one row per submitted feedback session
5. `Ratings` — one row per skill rating per session

---

## Tab 1: `Config`

Update this once a year when a new cohort starts.

| Year | Label | Active |
|------|-------|--------|
| 2025 | MAud 2025 | TRUE |
| 2026 | MAud 2026 | TRUE |
| 2024 | MAud 2024 | FALSE |

- **Year**: Cohort year identifier — used as a key everywhere
- **Label**: Display name in the dashboard cohort dropdown
- **Active**: Set to FALSE when a cohort graduates (data is kept but hidden from default view)

---

## Tab 2: `Students`

Add students here at the start of each year.

| StudentID | Cohort | Name | Email |
|-----------|--------|------|-------|
| STU2025-001 | 2025 | Alice Smith | asmith@auckland.ac.nz |
| STU2025-002 | 2025 | Ben Jones | bjones@auckland.ac.nz |
| STU2026-001 | 2026 | Cara Lee | clee@auckland.ac.nz |

- **StudentID**: Unique ID — suggested format: `STU{year}-{number}`
- **Cohort**: Must match a Year value in Config
- **Name**: Full name, shown in dashboard and on PDFs
- **Email**: Used for notification emails when a session is submitted

---

## Tab 3: `Skills`

Defines all skills and their S1/S2/Y2 progression targets. Set up once; rarely changed.

Columns: `SkillID | SkillName | Objective | Scope | ExpS1 | ExpS2 | ExpY2`

### Scope values

| Scope | Included in clinic types |
|-------|--------------------------|
| `all` | All five clinic types (including Other/observations) |
| `all-nonobs` | Adult Diagnostic, Paediatric Diagnostic, Adult Rehabilitation, Paediatric Rehabilitation |
| `all-dx` | Adult Diagnostic, Paediatric Diagnostic |
| `adult-dx` | Adult Diagnostic only |
| `paed-dx` | Paediatric Diagnostic only |
| `adult-rehab` | Adult Rehabilitation only |
| `paed-rehab` | Paediatric Rehabilitation only |

### General skills

| SkillID | SkillName | Objective | Scope | ExpS1 | ExpS2 | ExpY2 |
|---------|-----------|-----------|-------|-------|-------|-------|
| GEN-APPEAR | Appearance | Obj I A: Professional & Ethical Behaviour | all | 4 | 4 | 4 |
| GEN-SAFETY | Clinical Safety | Obj I A: Professional & Ethical Behaviour | all | 3 | 4 | 5 |
| GEN-CONFID | Confidentiality | Obj I A: Professional & Ethical Behaviour | all | 3 | 4 | 5 |
| GEN-INTERP | Interpersonal / Rapport | Obj I B: Interpersonal Behaviour | all | 2 | 3 | 4 |
| GEN-VOICE | Voice and Language | Obj I B: Interpersonal Behaviour | all | 2 | 3 | 4 |
| GEN-REFINE | Refining Behaviour | Obj I C: Refining Behaviour | all | 3 | 4 | 4 |
| SESS-PREP | Session Preparation | Obj II: Session Preparation | all | 2 | 4 | 4 |
| GEN-MAINT | Clinic Maintenance | Obj I A: Professional & Ethical Behaviour | all-nonobs | 2 | 4 | 5 |
| GEN-RECORD | Record Keeping | Obj I A: Professional & Ethical Behaviour | all-nonobs | 2 | 3 | 4 |
| SESS-TIME | Speed and Time | Obj III: Session Management | all-nonobs | 2 | 3 | 4 |
| REAS-INTER | Interpretation | Obj IV: Reasoning & Integration | all-nonobs | 1 | 2 | 4 |
| WRIT-PROG | Progress Notes | Obj V D: Clinical Writing | all-nonobs | 1 | 2 | 3 |
| WRIT-REPT | Report Writing | Obj V D: Clinical Writing | all-nonobs | 1 | 2 | 3 |
| PRES-EXPL | Explanation | Obj V C: Presentation of Findings | all-nonobs | 2 | 4 | 5 |
| PRES-RECO | Recommendations | Obj V C: Presentation of Findings | all-nonobs | 2 | 3 | 4 |
| REAS-XCHK | Applies X-check | Obj IV: Reasoning & Integration | all-dx | 2 | 3 | 4 |
| WRIT-AUDG | Audiogram | Obj V D: Clinical Writing | all-dx | 3 | 4 | 5 |

### Adult Diagnostic skills

| SkillID | SkillName | Objective | Scope | ExpS1 | ExpS2 | ExpY2 |
|---------|-----------|-----------|-------|-------|-------|-------|
| ADX-HIST | History Taking | Obj V A: Clinical Interview | adult-dx | 2 | 4 | 5 |
| ADX-OTOS | Otoscopy | Obj V B: Testing | adult-dx | 2 | 3 | 4 |
| ADX-TYMP | Tympanometry | Obj V B: Testing | adult-dx | 2 | 3 | 5 |
| ADX-ART | ARTs | Obj V B: Testing | adult-dx | 2 | 3 | 4 |
| ADX-PTA | PTA (AC/BC) | Obj V B: Testing | adult-dx | 2 | 3 | 4 |
| ADX-PTMSK | PT Masking | Obj V B: Testing | adult-dx | 3 | 5 | 5 |
| ADX-SPAUD | Speech Audiometry | Obj V B: Testing | adult-dx | 3 | 4 | 4 |
| ADX-SPMSK | Speech Masking | Obj V B: Testing | adult-dx | 2 | 3 | 4 |

### Paediatric Diagnostic skills

| SkillID | SkillName | Objective | Scope | ExpS1 | ExpS2 | ExpY2 |
|---------|-----------|-----------|-------|-------|-------|-------|
| PDX-HIST | History Taking | Obj V A: Clinical Interview | paed-dx | 1 | 2 | 2 |
| PDX-OTOS | Otoscopy | Obj V B: Testing | paed-dx | 1 | 2 | 3 |
| PDX-VRAD | VRA (Distractor) | Obj V B: Testing | paed-dx | 1 | 2 | 3 |
| PDX-VRAT | VRA (Tester) | Obj V B: Testing | paed-dx | 1 | 2 | 3 |
| PDX-PLAY | PLAY | Obj V B: Testing | paed-dx | 1 | 2 | 3 |
| PDX-KTT | KTT | Obj V B: Testing | paed-dx | 1 | 2 | 3 |
| PDX-OAE | OAEs | Obj V B: Testing | paed-dx | 1 | 3 | 4 |

### Adult Rehabilitation skills

| SkillID | SkillName | Objective | Scope | ExpS1 | ExpS2 | ExpY2 |
|---------|-----------|-----------|-------|-------|-------|-------|
| ARH-DISC | HA Discussion | Obj V A: Clinical Interview | adult-rehab | 1 | 2 | 3 |
| ARH-NEED | Needs Assessment (COSI/HHIA) | Obj V A: Clinical Interview | adult-rehab | 1 | 2 | 3 |
| ARH-FUND | Funding | Obj IV: Reasoning & Integration | adult-rehab | 1 | 2 | 3 |
| ARH-HSEL | HA Selection | Obj IV: Reasoning & Integration | adult-rehab | 1 | 2 | 3 |
| ARH-TRBL | Troubleshooting | Obj IV: Reasoning & Integration | adult-rehab | 1 | 2 | 3 |
| ARH-IMPR | Impressions | Obj V B: Testing | adult-rehab | 1 | 2 | 3 |
| ARH-NOAH | Noah / HA Setup | Obj V B: Testing | adult-rehab | 1 | 2 | 3 |
| ARH-PROG | HA Programming | Obj V B: Testing | adult-rehab | 1 | 2 | 3 |
| ARH-VERF | Verification | Obj V B: Testing | adult-rehab | 1 | 2 | 3 |
| ARH-VALD | Validation | Obj V B: Testing | adult-rehab | 1 | 2 | 3 |
| ARH-COUN | Counselling | Obj V C: Presentation of Findings | adult-rehab | 1 | 2 | 3 |

### Paediatric Rehabilitation skills

| SkillID | SkillName | Objective | Scope | ExpS1 | ExpS2 | ExpY2 |
|---------|-----------|-----------|-------|-------|-------|-------|
| PRH-NEED | Needs Assessment | Obj V A: Clinical Interview | paed-rehab | 1 | 1 | 2 |
| PRH-HSEL | HA Selection | Obj IV: Reasoning & Integration | paed-rehab | 1 | 1 | 2 |
| PRH-PROG | HA Programming | Obj V B: Testing | paed-rehab | 1 | 1 | 2 |
| PRH-VERF | Verification | Obj V B: Testing | paed-rehab | 1 | 1 | 3 |
| PRH-VALD | Validation | Obj V B: Testing | paed-rehab | 1 | 1 | 2 |
| PRH-COUN | Counselling | Obj V C: Presentation of Findings | paed-rehab | 1 | 1 | 2 |

---

## Tab: `Supervisors`

| Name | Email | IsCoordinator |
|------|-------|----------------|
| Dr Smith | dsmith@auckland.ac.nz | FALSE |
| Dr Jones | djones@auckland.ac.nz | TRUE |

- **Name**: Must match the free-text name typed into Supervisor 1/2 on the feedback form (case-insensitive) so the session PDF email can be routed to them
- **Email**: Where session PDFs are sent
- **IsCoordinator**: TRUE/FALSE — coordinators receive a copy of every session PDF (supervisor feedback *and* student-logged hours), not just sessions they personally supervised

---

## Tab 4: `Sessions`

One row per submitted feedback form. The form writes here via Apps Script.

| Column | Description |
|--------|-------------|
| Timestamp | Auto-set when submitted |
| SessionID | Unique ID generated on submission (e.g. `SES-20250315-001`) |
| StudentID | Must match Students tab |
| Date | Session date (YYYY-MM-DD) |
| Supervisor1 | Primary supervisor name |
| Supervisor2 | Second supervisor name (may be blank) |
| Location | Clinic location |
| Activities | Comma-separated list of checked activities |
| Hrs_AdultDx_Obs | Adult Diagnostic — observed hours |
| Hrs_AdultDx_Test | Adult Diagnostic — tested/direct hours |
| Hrs_PaedDx_Obs | Paediatric Diagnostic — observed hours |
| Hrs_PaedDx_Test | Paediatric Diagnostic — tested/direct hours |
| Hrs_AdultRehab_Obs | Adult Rehabilitation — observed hours |
| Hrs_AdultRehab_Test | Adult Rehabilitation — tested/direct hours |
| Hrs_PaedRehab_Obs | Paediatric Rehabilitation — observed hours |
| Hrs_PaedRehab_Test | Paediatric Rehabilitation — tested/direct hours |
| Hrs_Other_Obs | Other — observed hours |
| Hrs_Other_Test | Other — tested/direct hours |
| Hrs_ORL | ORL hours (separate count) |
| Hrs_SLT | SLT hours (separate count) |
| Hrs_Simulation | Simulation hours (counts toward Other but tracked separately) |
| Hrs_Supervision | Clinical supervision hours (separate count) |
| Feedback_Well | "Done well" free text |
| Feedback_Improve | "To improve" free text |
| Feedback_General | General comments free text |

---

## Tab 5: `Ratings`

One row per skill rating per session. The form writes here once per rated skill.

| Column | Description |
|--------|-------------|
| Timestamp | Auto-set when submitted |
| SessionID | Links to Sessions tab |
| StudentID | Denormalised for easier querying |
| ClinicType | Exact value: `Adult Diagnostic`, `Paediatric Diagnostic`, `Adult Rehabilitation`, `Paediatric Rehabilitation`, or `Other` |
| Milestone | `S1`, `S2`, or `Y2` — the current progression milestone |
| SkillID | Must match Skills tab |
| Rating | Integer 1–5 |
| IsPriority | TRUE or FALSE |
| IsStrength | TRUE or FALSE |
| Comment | Supervisor comment (only present when IsPriority or IsStrength is TRUE) |

Example rows:

| Timestamp | SessionID | StudentID | ClinicType | Milestone | SkillID | Rating | IsPriority | IsStrength | Comment |
|-----------|-----------|-----------|------------|-----------|---------|--------|------------|------------|---------|
| 2025-03-15 14:30 | SES-20250315-001 | STU2025-001 | Adult Diagnostic | S1 | ADX-PTA | 3 | FALSE | TRUE | Good technique with masking |
| 2025-03-15 14:30 | SES-20250315-001 | STU2025-001 | Adult Diagnostic | S1 | ADX-PTMSK | 2 | TRUE | FALSE | Needs to practise identifying when masking is needed |

- The dashboard uses the most recent rating per skill per milestone
- Only rows where IsPriority or IsStrength is TRUE will have a Comment
- Numeric ratings do **not** appear on student PDFs — only flagged skills with comments

---

## Rating scale

| Value | Label | Description |
|-------|-------|-------------|
| 1 | Absent | Unsafe or unable to perform the task without specific instruction |
| 2 | Emerging | Can attempt task; requires significant hand-holding, significant issues around timing |
| 3 | Present | Can perform for routine cases; minor concerns, generally minimal supervisor input, may lack confidence |
| 4 | Developed | Appears confident, manages time effectively; needs assistance in complex cases only |
| 5 | Consistent | Manages complex clients well; high level of competence |

---

## Progression milestones

| Code | Name | When |
|------|------|------|
| S1 | Semester 1 | End of first semester, Year 1 |
| S2 | Semester 2 | End of second semester, Year 1 |
| Y2 | End of Year 2 | End-of-course goal |

---

## Yearly workflow

**New cohort:**
1. Config tab → add row: `2027 | MAud 2027 | TRUE`
2. Students tab → add students with cohort `2027`
3. Dashboard automatically shows the new cohort — no other changes needed

**Cohort graduates:**
1. Config tab → set their Active to `FALSE`
2. Data is retained; cohort is hidden from the default dashboard view

---

## Deployment

1. In Google Sheets: **Extensions > Apps Script**
2. Paste the contents of `src/apps-script/Code.gs`
3. Also paste `src/apps-script/PlacementImporter.gs` if using the placement hours importer
4. **Deploy > New deployment > Web app**
   - Execute as: **Me**
   - Who has access: **Anyone within University of Auckland** (or "Anyone" if supervisors access from outside)
5. Copy the web app URL
6. Paste it into `src/dashboard/index.html` and `src/form/index.html` at the `API_URL` constant (currently set to `null` for demo mode)
