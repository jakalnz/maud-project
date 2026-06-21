# PebblePad Build — Session Handover (June 2026)

This picks up from `HANDOVER_PebblePad.md` and `MAud_PebblePad_Workbook_Spec.docx`. Everything below was worked out in a design conversation, not yet built in PebblePad unless noted. Treat this as the current source of truth where it conflicts with the original spec doc — several things changed after testing the live system.

---

## Key mechanism: Collection + tagged Template, not native workbook pages

The workbook's repeating "pages" (Page 2, Page 3) are **not** built as pages literally inside the workbook. The actual pattern:

1. A **standalone Template** (separate Resource) holds the real form/fields.
2. A **Collection placeholder page** inside the workbook aggregates instances by tag.
3. The Collection page has a hyperlinked "Use this template" instruction pointing at the Template.
4. Cascading tags on both the Collection's criteria and the Template's properties keep everything wired together automatically.

Student experience: open the Collection page in their workbook → click the link → fill in a fresh instance of the Template → save → it appears in the Collection (assuming tags + auto-search are configured correctly).

**Known footgun, already hit once:** the Collection's criteria has a setting called "Automatically turn on search for new responses." If it's not ticked, the Collection's list doesn't visibly update even though the underlying tagged asset exists — this is what caused the "nothing showing up" issue on Page 2 early on. Tick it on every new Collection from the start.

---

## Page 2 — Session reflection (current state)

**Status:** live, in use by current cohort. Built before most of the patterns below were confirmed, so it has some inconsistencies worth fixing.

- **Save-as title:** `[Session ID] ([Supervisor]) "[Title]"`. **Confirmed:** the `SES-` and `STU-` prefixes are both correct, not a mismatch — `SES-` is generated when a supervisor submits the feedback form, `STU-` when a student self-logs hours via the "Log Supervised Hours" page. This is a useful signal, not noise: it tells you at a glance whether an entry originated from the supervisor or the student. See the dashboard parked item below — this prefix is the natural basis for the internal/external differentiation feature.
- **Clinic type:** changed from single dropdown to **multi-select checkboxes** — sessions can span multiple types, and "Other" needed real subcategories (ORL Observation, SLT Observation, Clinical Supervision, Other–specify). This field is informational only; it doesn't need to match the backend's exact 5-string clinic-type taxonomy, since that precision already lives in the Sessions sheet via the supervisor form or student hours-logging form.
- **Feedback/evidence upload:** one media picker block, doing double duty — holds either the supervisor's rated feedback PDF (internal sessions) or the signed hours-verification PDF (external placements). Same template is reused for both cases rather than building two separate page types.
- **Capability self-assessment (R/A/G per objective): REMOVED.** Decision made this session — not making students do per-session self-ratings. This concept is being folded into the parked "calibration exercise" item (see Parked Items below), likely living in the web dashboard rather than PebblePad. **Action needed:** if this was ever built into the live page, remove it; sections after it shift up one letter.
- **Reflection prompts: confirmed staying** as the Borton/Driscoll "What? / So What? / Now what?" cycle (3 prompts). The original spec's four scaffolded prompts (responding to priorities / building on strengths / overall learning / goals) were discussed as an alternative early on but are not being adopted — no further action needed here.
- **Sharing instruction (new):** hint text added near the save block:
  > After saving this reflection, if it's for an internal clinic session, share it with your supervisor: open it from your Asset Store, select "I want to... → Share → With people," search for your supervisor by name or university email, and set access to "View" (you can allow comments too). If this is an external placement (e.g. ORL or Speech-Language observation), share it with your course coordinator instead, so they know the reflection is ready before they confirm your hours.
- **Page Verification:** the "Lock page on verification" property **exists and is confirmed available** on this template (checked in Resource Builder). **Not yet tested end-to-end** — needs a real second account (a friendly student) to confirm the full verify-and-lock loop actually works when the asset is surfaced via a Collection rather than a native workbook page. Test plan already written; just needs a volunteer.

---

## Page 3 — Term summary (newly designed this session, not yet built)

Repeating, one per milestone (S1 / S2 / Y2) — same Collection + Template pattern as Page 2, but as its **own separate Collection** with its own tag (e.g. `term-summary`), distinct from Page 2's tag.

**Save-as title:** `[Milestone] — Term Summary` (e.g. "S1 — Term Summary")

**Top of page:** Milestone dropdown (S1 / S2 / Y2), required. Drives the title and gives a structured field for "which term" without needing a separate manual tag.

**Section A — Term overview**
- One rich text block, hint text along the lines of: reflect on patterns across the term, which skills improved most, which are still sticking points, **at least** 200 words (it's a floor, not a target — word it as "at least 200" not "aim for 200").
- **Confirmed:** minimum word counts can be enforced as a real validation rule on the block, not just advisory hint text — set the actual minimum on the block itself, not only in the wording.

**Section B — Hours progress: REMOVED.** Decided this session — not needed. (Originally: total hours logged, progress against compliance targets, plan for any behind category.)

**Section C — Full capability self-assessment.** Confirmed staying, as designed:
- **Five separate capability blocks** (not the four-block Obj II/III combined grouping originally discussed) — Obj I, II, III, IV, V, each its own block. This was a deliberate split decided this session, matching the Guidelines document's five distinct objectives.
- **Numeric capability sub-type, 5-point scale (1–5), not R/A/G text on the block itself.** Confirmed in testing: the Numeric capability block's point labels can't be relabelled to custom text (Red/Amber/Green) — they're numeric only. Decision: use the same 1–5 scale already used everywhere else in the system (Absent/Emerging/Present/Developed/Consistent), rather than forcing a 3-point R/A/G scale that would be inconsistent with the rest of the project. One shared instruction line above all five blocks explains the 1–5 scale once, rather than repeating it per block.
- **R/A/G still exists, but at a different layer:** Capability Approval statements (configured in ATLAS — Management → Feedback → Capability settings) let an assessor stamp a Red/Amber/Green verdict *on top of* a student's numeric self-rating, when reviewing in Submission Viewer. This is independent of the student-facing scale. Walkthrough for enabling it is below.
- **Open question raised but not resolved:** Page 2 never had a capability section in the first place (removed, see above), so there's no four-vs-five-block inconsistency to worry about between the two pages anymore.

**Section D — Goal setting.** No true repeating sub-block exists in PebblePad for "add another goal" within one page. Built as **three fixed slots** (Goal 1 required, Goals 2–3 optional), each with: Objective dropdown (now I/II/III/IV/V, five options matching Section C), "what does success look like" text area, "steps to get there" text area. **Unconfirmed:** whether "required" is a true per-block property or could end up page-wide — check this doesn't accidentally force Goals 2/3 as well.

**Section E — Evidence.** Was Section E in the original spec; now Section D after Section B's removal. Drop the "calibration reports" example from the hint text — that feature doesn't exist yet, no point referencing it and confusing students; it costs nothing to add back in once calibration ships, since editing a resource updates all existing instances. **Unconfirmed:** whether a single media picker block accepts multiple attached assets or just one — if just one, use 2–3 fixed optional slots (same pattern as Goals) instead of a single block.

**Sharing instruction:** same "share with people" pattern as Page 2, but no supervisor involved at the term level — not yet drafted since Page 3 doesn't have an obvious individual recipient the way Page 2 does (term summaries go to the coordinator via the Assignment/Submission Viewer route below, not a manual share).

---

## Page 3's submission routing — different from Page 2

Because Term Summaries are infrequent (3 per student across the whole programme, vs. potentially dozens of session reflections), the decision was made to route Page 3 through a **dedicated ATLAS Assignment** rather than relying purely on the Collection for visibility. This is what makes Capability Approval and Page Verification actually usable on this page — those assessor tools only attach to assets properly submitted to an Assignment, not to assets that are merely tag-linked into a Collection for display.

**Setup steps (not yet done):**
1. Management → Assignments → create a new assignment (e.g. "Term Summary"), separate from whatever governs the main workbook.
2. Workspace → Resources tab → Add a resource → add the Term Summary template directly (separate step from it being linked into the Collection — that link is for student-facing navigation only, this step is what connects it to ATLAS).
3. On that resource's edit/cog icon → "Would you like to prompt the user to auto-submit when saved?" → Yes → select the Term Summary assignment.

**Flagged but unconfirmed:** this combination — a template that's both tag-linked into a Collection *and* separately added to the workspace for auto-submit — hasn't been seen directly documented anywhere. No obvious reason the two mechanisms would conflict (they're built on different parts of PebblePad), but it hasn't been tested. Test before trusting it for the live cohort.

---

## Capability Approval setup walkthrough (for when Page 3 is built)

1. ATLAS workspace → Management → Feedback → Capability settings tab.
2. Enable capability approvals → Yes.
3. Enter the three statements (from the original spec):
   - Red: "Below expected progression — significant concerns or not yet attempted. Action plan required."
   - Amber: "Approaching expected progression — developing but not yet meeting term expectations."
   - Green: "Meeting or exceeding expected progression — on track for this stage of the programme."
4. Save. This is workspace-wide, one-time — covers every capability block in every submission in the workspace, both pages.
5. Management → Managers → find the relevant manager role → Modify Permissions → tick "Add capability approval" → Save.
6. **Important:** with Sets dropped and individual session supervisors moved to lightweight "share with people" access (see below), they are not Managers and can't see Submission Viewer at all. In the current design, this permission is only meaningful for the course coordinator (Lead Tutor). Decide if it's worth enabling given that limited audience, or leave switched off.

---

## Page Verification setup (confirmed available, not yet tested end-to-end)

- Page Properties → Template tab → Page locking section → tick "Lock page on verification."
- Later, in Submission Viewer: "I want to..." → "Verify this page" — confirms completion and locks the page if the property above is ticked.
- Confirmed: this requires the reviewer to be a full Manager/Lead Tutor, not an External Assessor account (External Assessors can't verify pages).
- **Still needs a real test** with a second account (friendly student), since this asset is surfaced via a Collection rather than a native workbook page, and that combination hasn't been directly confirmed elsewhere. A quick partial check possible without a second account: open any existing real submission you already have manager access to, and check whether "Verify this page" shows up as an option at all — doesn't test the lock, but narrows things.

---

## Routing decision: dropped Sets, using "Share with people" instead

**Original problem:** wanted students to route reflections to the specific supervisor they were with, falling back to the course coordinator. ATLAS "Sets" can restrict a Manager's visibility to a sub-group of students, but Sets group at the **student level**, persistently — doesn't map cleanly onto supervision that rotates per session/block, and requires ongoing manual maintenance as rotations change.

**Decision: dropped Sets entirely.** Using PebblePad's separate, independent **"Share with people"** mechanism instead:
- Student selects one specific asset (one reflection, not the whole workbook) → "I want to..." → Share → "With people" → search recipient by name/email → set permission (View, optionally allow comments) → share.
- This is completely outside the ATLAS/workspace/Manager machinery — no Sets, no group maintenance, exact per-entry precision.
- **Tradeoff, accepted knowingly:** this is a manual step the student has to remember each time (mitigated by putting it directly in the page's hint text). It also likely gives view/comment access only, not the full assessor toolset — judged acceptable since individual session supervisors aren't doing formal sign-off on session reflections, only the course coordinator does (via the Assignment/Submission Viewer route for Term Summaries and external placement hours-verification).
- **Role split this produces:** the course coordinator stays a Manager/Lead Tutor (needed for Capability Approval + Page Verification). Individual session supervisors likely don't need Manager/workspace enrolment at all anymore — just a basic PebblePad account, which UoA staff have by default.

---

## Submission model: one-off vs continuous (worth revisiting, not yet acted on)

Discovered that the live workbook is set up with a **one-off/periodic submission** (student manually "submits the whole workbook," ATLAS gets a frozen duplicate copy at that moment) rather than **"Modify until deadline"** (continuous sync). This was flagged as a likely cause of visibility problems, but turned out to be less urgent than first thought: testing showed the **Collection's contents update independently of the workbook-level freeze** — new tagged reflections appear in the Collection's underlying data regardless of submission state. The actual blocker was the "automatically turn on search for new responses" toggle (see top of doc), not the submission model.

**Net effect:** the auto-submit / continuous-deadline question is now lower priority than originally thought, but not fully resolved either. If visibility issues recur after fixing the auto-search toggle, revisit:
- Resource-level "auto-submit when saved" toggle (Workspace → Resources → edit icon).
- Assignment-level "Modify until deadline" vs one-off setting (Management → Assignments).

---

## Dashboard confirmation step (Apps Script / Sheets system, not PebblePad)

New requirement surfaced this session, applies to **all hours, internal and external**:

- Clinic supervisors need to go into the existing skills/hours dashboard and use the **already-existing "Confirm" button** against a session's hours — but now, this should happen *after* they've read the student's shared PebblePad reflection, since the confirm action doubles as the coordinator's signal that a reflection has been completed.
- **Future dashboard work (explicitly deferred to a later Claude Code session, not now):** a way to differentiate internal vs external hours in the coordinator's dashboard view, and a filter for confirmed/not-confirmed entries, so the coordinator can quickly spot what still needs attention.
- For external placements specifically, the course coordinator does the confirming (no individual clinic supervisor is involved in that workflow) — consistent with their existing Page Verification / hours-review role.

---

## Communications drafted (see `MAud_PebblePad_Comms_Pack.docx`)

Two explainer emails (students, supervisors), a process diagram, and Outlook auto-folder setup instructions — all saved as a single document, ready to copy into actual emails. Content reflects every decision above (sharing pattern, dropped Sets, dashboard confirmation step). Diagram source is also available as `workflow_diagram.svg` if it needs editing later.

---

## Parked / deferred items (updated)

Carried over from the original handover, plus new additions from this session:

- **Self-skill-assessment + calibration comparison** — merged into one parked item. Originally "blind self-rating comparison on dashboard generating a downloadable evidence report for PebblePad." Page 2's capability self-assessment was dropped specifically because this functionality is meant to live here instead, likely in the web app, not duplicated in PebblePad. Page 3's capability section is unaffected — that one stays in PebblePad since it's tied to formal end-of-term Capability Approval.
- **Dashboard: internal/external differentiation + confirmed/unconfirmed filter** — needed to support the dashboard confirmation step above. Build in Claude Code, not now, but the implementation path is clearer: the `SES-`/`STU-` ID prefix (confirmed above) already distinguishes supervisor-submitted from student-self-logged entries, which lines up closely with internal vs external — likely the simplest basis for the differentiation rather than building a new flag from scratch.
- Extended milestone structure beyond S1/S2/Y2 (Summer Block, T5–T8).
- Automated placement competency mapping.
- PebblePad API push / Dashboard embed in PebblePad iframe — both still blocked on UoA IT/PebblePad admin action (whitelist request, API client credentials).
- A possible third explainer email for the course coordinator, covering both their dashboard-confirm duty and PebblePad page verification duty in one place — raised but not drafted; coordinator audience is small enough that this might just be a direct conversation instead.

---

## Things still genuinely uncertain (don't assume, test)

- Whether "required" on a block is truly per-block or could be page-wide.
- Whether a single media picker block holds multiple assets.
- Whether Page Verification + Collection-surfaced assets actually work together end to end.
- Whether the Collection-tag + Assignment-auto-submit combination (Page 3's routing) works without conflict.
