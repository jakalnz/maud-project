# Google Sheets Setup Guide

## Sheet Structure

Create a new Google Sheet with **four tabs** (sheets):

---

### Tab 1: `Config`

This is what you update once a year when a new cohort starts.

| Year | Label | Active |
|------|-------|--------|
| 2025 | MAud 2025 | TRUE |
| 2026 | MAud 2026 | TRUE |
| 2024 | MAud 2024 | FALSE |

- **Year**: The cohort year identifier (used as a key everywhere)
- **Label**: Display name shown in the dashboard dropdown
- **Active**: Set to FALSE when a cohort has graduated (hides from default view but data remains)

---

### Tab 2: `Students`

Add new students here at the start of each year.

| StudentID | Cohort | Name | Email |
|-----------|--------|------|-------|
| STU2025-001 | 2025 | Alice Smith | asmith@auckland.ac.nz |
| STU2025-002 | 2025 | Ben Jones | bjones@auckland.ac.nz |
| STU2026-001 | 2026 | Cara Lee | clee@auckland.ac.nz |

- **StudentID**: Unique ID (suggest format: STU{year}-{number})
- **Cohort**: Must match a Year value in the Config tab
- **Name**: Student's name (displayed in dashboard)
- **Email**: Optional, for future notification features

---

### Tab 3: `Skills`

This defines all the skills and their expected progression. You set this up once and rarely change it.

| SkillID | SkillName | Objective | ClinicTypes | ExpT1 | ExpT2 | ExpT3 | ExpT4 |
|---------|-----------|-----------|-------------|-------|-------|-------|-------|
| GEN-APPEAR | Appearance | Obj I: Professional behaviour | All | 5 | 5 | 5 | 5 |
| GEN-MAINT | Clinic maintenance | Obj I: Professional behaviour | All | 3 | 5 | 5 | 5 |
| GEN-RECORD | Record keeping | Obj I: Professional behaviour | All | 2 | 3 | 4 | 5 |
| GEN-SAFETY | Clinical safety | Obj I: Professional behaviour | All | 2 | 3 | 5 | 5 |
| GEN-CONFID | Confidentiality | Obj I: Professional behaviour | All | 3 | 3 | 5 | 5 |
| GEN-INTERP | Interpersonal | Obj I: Professional behaviour | All | 2 | 3 | 4 | 5 |
| GEN-VOICE | Voice and language | Obj I: Professional behaviour | All | 2 | 3 | 4 | 5 |
| GEN-REFINE | Refining behaviour | Obj I: Professional behaviour | All | 3 | 5 | 5 | 5 |
| SESS-PREP | Session preparation | Obj II/III: Preparation and management | All | 3 | 4 | 5 | 5 |
| SESS-TIME | Speed and time | Obj II/III: Preparation and management | All | 2 | 3 | 4 | 5 |
| ADX-HIST | History taking | Obj V B: Testing | Adult diagnostic | 2 | 3 | 4 | 5 |
| ADX-OTOS | Otoscopy | Obj V B: Testing | Adult diagnostic | 2 | 3 | 4 | 5 |
| ADX-TYMP | Tympanometry | Obj V B: Testing | Adult diagnostic | 1 | 2 | 4 | 5 |
| ADX-ART | Acoustic reflexes | Obj V B: Testing | Adult diagnostic | 1 | 2 | 3 | 5 |
| ADX-PTA | PTA (AC/BC) | Obj V B: Testing | Adult diagnostic | 3 | 3 | 4 | 5 |
| ADX-MASK | ID need to mask | Obj V B: Testing | Adult diagnostic | 3 | 3 | 4 | 5 |
| ADX-PTMSK | PT masking | Obj V B: Testing | Adult diagnostic | 2 | 3 | 3 | 4 |
| ADX-SPAUD | Speech audiometry | Obj V B: Testing | Adult diagnostic | 2 | 3 | 4 | 5 |
| ADX-SPMSK | Speech masking | Obj V B: Testing | Adult diagnostic | 2 | 2 | 3 | 4 |
| REAS-INTER | Interpretation | Obj IV: Reasoning and integration | Adult diagnostic | 2 | 3 | 4 | 4 |
| REAS-XCHK | Applies X-check | Obj IV: Reasoning and integration | Adult diagnostic | 2 | 3 | 4 | 5 |
| PRES-EXPL | Explanation | Obj V C: Presentation of findings | All | 2 | 3 | 4 | 5 |
| PRES-RECO | Recommendations | Obj V C: Presentation of findings | All | 1 | 2 | 3 | 3 |
| WRIT-PROG | Progress notes | Obj V D: Clinical writing | All | 1 | 1 | 3 | 4 |
| WRIT-REPT | Report writing | Obj V D: Clinical writing | All | 1 | 1 | 2 | 3 |
| WRIT-AUDG | Audiogram | Obj V D: Clinical writing | All | 3 | 4 | 5 | 5 |
| PAED-HIST | History taking | Obj V B: Testing | Paediatric | 1 | 1 | 2 | 2 |
| PAED-VRAD | VRA (distractor) | Obj V B: Testing | Paediatric | 1 | 1 | 1 | 2 |
| PAED-VRAT | VRA (tester) | Obj V B: Testing | Paediatric | 1 | 1 | 1 | 2 |
| PAED-PLAY | Play | Obj V B: Testing | Paediatric | 1 | 1 | 1 | 2 |
| PAED-KTT | KTT | Obj V B: Testing | Paediatric | 1 | 1 | 1 | 2 |
| PAED-OAE | OAEs | Obj V B: Testing | Paediatric | 1 | 1 | 2 | 3 |
| REHAB-DISC | HA discussion | Obj V B: Testing | Adult rehab | 1 | 1 | 1 | 2 |
| REHAB-NEED | Needs assessment | Obj V B: Testing | Adult rehab | 1 | 1 | 1 | 2 |
| REHAB-FUND | Funding | Obj IV: Reasoning and integration | Adult rehab | 1 | 1 | 1 | 2 |
| REHAB-HSEL | HA selection | Obj IV: Reasoning and integration | Adult rehab | 1 | 1 | 1 | 2 |
| REHAB-TRBL | Troubleshooting | Obj IV: Reasoning and integration | Adult rehab | 1 | 1 | 1 | 2 |
| REHAB-IMPR | Impressions | Obj V B: Testing | Adult rehab | 1 | 1 | 1 | 2 |
| REHAB-NOAH | Noah/HA setup | Obj V B: Testing | Adult rehab | 1 | 1 | 1 | 2 |
| REHAB-PROG | HA programming | Obj V B: Testing | Adult rehab | 1 | 1 | 1 | 2 |
| REHAB-VERF | Verification | Obj V B: Testing | Adult rehab | 1 | 1 | 1 | 2 |
| REHAB-VALD | Validation | Obj V B: Testing | Adult rehab | 1 | 1 | 1 | 2 |
| REHAB-COUN | Counselling | Obj V C: Presentation of findings | Adult rehab | 1 | 1 | 1 | 2 |

- **ClinicTypes**: Comma-separated list. Use "All" for skills that appear in every clinic type.
- **ExpT1–T4**: Expected rating (1–5) at end of each term.

---

### Tab 4: `Ratings`

This is where assessment data accumulates. Each row is one skill rating from one session.

| Timestamp | StudentID | Supervisor | ClinicType | Term | SkillID | Rating | IsPriority | IsStrength | Comment |
|-----------|-----------|------------|------------|------|---------|--------|------------|------------|---------|
| 2025-03-15 14:30 | STU2025-001 | Dr Patel | Adult diagnostic | 1 | ADX-PTA | 3 | FALSE | TRUE | Good technique |
| 2025-03-15 14:30 | STU2025-001 | Dr Patel | Adult diagnostic | 1 | ADX-PTMSK | 2 | TRUE | FALSE | Needs to practise identifying when masking is needed |

- The feedback form writes to this tab (via the Apps Script `addRating` function)
- Multiple ratings per student per skill are fine — the dashboard uses the most recent per term
- **IsPriority / IsStrength**: The flags from the feedback form

---

## Yearly Workflow

When a new cohort starts:

1. Open the **Config** tab, add a new row: `2027 | MAud 2027 | TRUE`
2. Open the **Students** tab, add the new students with the 2027 cohort value
3. That's it — the dashboard will automatically show the new cohort

When a cohort graduates:

1. Set their **Active** column to `FALSE` in Config
2. Their data remains in the sheet but they're hidden from the default dashboard view
3. You can still access them by selecting "Show archived" in the dashboard

---

## Deployment

1. In Google Sheets: Extensions > Apps Script
2. Paste the contents of `Code.gs`
3. Deploy > New deployment > Web app
4. Set access to "Anyone within University of Auckland" (or "Anyone" if needed externally)
5. Copy the web app URL
6. Paste it into `dashboard.html` at the `API_URL` constant
7. Upload `dashboard.html` to your GitHub Pages repo
