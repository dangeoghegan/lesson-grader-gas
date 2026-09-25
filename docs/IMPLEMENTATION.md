# Lesson Grader — implementation & safe test plan

**Status (25 September 2026):** Repository implementation and mock-tested UI. **Not deployed or validated against Dan's live Apps Script, Drive, Sheets, Classroom or Gemini account.** `docs/AUDIT.md` documents the original repository at commit `a468449`; its findings remain a baseline, not a description of the repaired files.

## What changed

- `Code.gs` now parses: removed the duplicated, stale second copy and the dangling block at the seam. The original Sheets dialog is retained as a **Jewellery-only legacy workflow**, with explicit guards to prevent grading other projects with its 12 hardcoded criteria. Corrected several header-based approval/lock writes, task-scoped import matching and version status handling. The two legacy AI entry points now share one ticking policy, filter invalid observable IDs and record successful evidence proposals. The legacy classroom menu still exists; the new web workflow does **not** silently use it.
- `RubricEngine.gs` builds **task-specific reviewed rubric profiles** from Shared Rubric spreadsheet rows or manually entered rows. Profiles support any project with named criteria, marks and observable descriptors. A–E weights come from the workbook's `Setup` sheet (or documented defaults). Other band names are allowed **only with a teacher-provided 0–1 weight**. Missing marks and unknown band weights block activation. Gemini's extraction is a *draft*, never a grade or automatically activated rubric.
- `WebGrading.gs` provides workbook-scoped APIs for task dashboards, **explicit backup-first enablement** of grading tables, Classroom assignment import, rubric activation, per-student grading, AI *proposal*, teacher draft, explicit approval, reassessment and final PDF. The server, not the browser or AI, computes official scores; multi-band evidence is capped at a criterion's maximum. Each submitted draft is bound to an immutable `RubricProfileID`; approvals append to `ApprovedGrades` and `AssessmentHistory`.
- `WebApp.html` is a responsive teacher workspace (desktop sidebar / mobile bottom navigation) with dashboard, task roster, evidence, scoring studio, feedback, rubric builder/library, Classroom link and Gemini Settings. Mic dictation in the web studio is opt-in and browser-dependent; speech goes into the teacher note, never directly into a grade.
- `WebAccess.gs` checks the accessing user's email against `LESSON_GRADER_ALLOWED_EMAILS` (missing configuration **denies access**). The Gemini key is kept only in Script Properties; never in the repository, a URL query string or an API response. Settings has **Save key** and **Test connection** actions. The test sends only `Ping`, not student data.
- `appsscript.json` declares the Classroom advanced service and V8 runtime; Node syntax/check/test harness and mock-only UI preview were added. **No credentials, student data or real Drive IDs are committed.**

## Before touching a live project

1. Export Dan's actual Apps Script project and **diff each file against this repository**. The original committed `Code.gs` did not parse, so a live project must already differ. Resolve any feature/data divergence; do not replace a live project wholesale.
2. Export the bound grading workbook and take a Drive copy. Record the file IDs, version and trigger/deployment settings. Check existing `Submissions`, `CriterionAssessments`, `ClassroomConfig`, `RubricProfiles` and `9DAT1 Marking` headers before deployment. Take separate backups of existing workbooks before using **Enable web grading**.
3. Confirm the institution allows student files to be sent to Gemini, who holds API/Drive/Classroom scopes, and whether the connected account can access every workbook. Do not test AI on student work without authorisation.
4. In the Apps Script editor, add/update each `.gs` and `.html` file plus `appsscript.json`; keep exactly **one** `doGet` (in `WebAppIntegration.gs`). Enable the **Google Classroom API** advanced service and Cloud API. The manifest deliberately leaves OAuth scope inference to Apps Script; review the consent screen when deploying.
5. In **Project Settings → Script Properties**, set `LESSON_GRADER_ALLOWED_EMAILS` to a comma-separated list of exact school teacher emails. Never paste this property's value into code. Email must be available from `Session.getActiveUser().getEmail()`. If a deployment returns a blank email, it fails closed; adjust deployment identity, not the auth check. For workbook creation, also set `LESSON_GRADER_WEB_CATEGORY` to the **existing** category folder name used inside each course's `Graded Assessments/Stage N/<course>/` path when there is more than one possible category; the app refuses to guess.
6. Deploy the web app **as the user accessing it**, restricted to the school/domain or narrower. **Do not select “Anyone” + “Execute as me”.** The allowed teacher must also have Drive/Sheets/Classroom permissions. Test with one allowed account and one denied account before inviting colleagues.
7. Open **Settings → Gemini API key** in the *real*, restricted deployment (not the mock preview), save a testing key and run **Test connection**. Or set `GEMINI_API_KEY` directly in Project Settings. This is a project-scoped secret shared by allowed teachers; rotate via the provider and replace it in Settings if exposed. No key is supplied in this repo; **a live connection cannot be verified here**. Network requests use the `x-goog-api-key` header. The current configuration tries `gemini-2.5-flash` then `gemini-2.5-flash-lite`; verify that those model IDs are available for the testing account before trialling student data.

The existing bound Sheets dialogs now also require an allowed teacher account for administrative/API operations. Configure the property **before** invoking those menu actions. A simple access-denied message is intentional; don't weaken it to make testing easier.

## Teacher flow: any project with a rubric

1. **Create workbook** for a course/year/task; optionally select a Shared Rubric file. Or open an existing initialised workbook. For an existing workbook, click **Create backup & enable** in the task dashboard before grading (additive column migration only). A selected Shared Rubric is a *source*, not automatically active.
2. Go to **Rubric library → Build rubric** or choose a Shared Rubric and preview its criteria, descriptions, band labels and marks. Alternatively upload PDF/Markdown/TXT (max 8 MB); Gemini extracts it into a new editable Shared Rubric source. Review against the original — AI can misread or omit marks. Enter missing marks, map non-A–E bands to explicit 0–1 weights, then tick the confirmation and click **Activate**. Without a reviewed profile, web grading refuses to proceed; it **never falls back to Jewellery**. Activating a revision appends a new `RubricProfiles` row and switches the task's active version. Existing assessment rows stay pinned to their original version.
3. **Classroom → Link** a course assignment to that workbook task. From the task dashboard, **Import from Classroom** (teacher action). Imports match `CourseID + CourseWorkID + StudentUserID`; changed attachment evidence creates a **new submission version** rather than overwriting an assessment. **No Classroom task?** Choose **Add student** to create a manual submission in this task's Sheets workbook; the student's work may be assessed from teacher-observed evidence and notes, but there is no Drive attachment or Classroom sync unless linked separately. Previously imported rows whose `Task` contains only an assignment title are *not* guessed or silently reassigned; reconcile them with a backup before using the web roster.
4. Select a student → review Drive links, grade-band descriptors, teacher notes and feedback. Tick supported evidence; the browser shows an **indicative** score, while server-side scoring is authoritative on save. **Save draft** logs a revision and does not publish grades. Unsaved edits are warned about. Optionally **Draft with Gemini** (max five eligible files; files are sent to Gemini) to get checkbox suggestions and evidence notes. Review every tick. AI never approves, locks or sends a grade to Classroom.
5. **Approve grade** only after a saved draft covers *all* criteria with evidence. The app writes a snapshot to the workbook's `ApprovedGrades` sheet and appends `AssessmentHistory`. **No change is made to the legacy `9DAT1 Marking` sheet or Google Classroom grade**. Approved/locked assessments are read-only. Generate a final PDF after approval. For changes, explicitly create a new reassessment version; older rows and their rubric versions remain available.

### Data stored in Google Sheets

| Tab | Role |
|---|---|
| `RubricProfiles` | Additive migration of original four-column sheet. New profile JSON, task name, source, schema version and active flag are header-mapped, not positional. Old JSON rows are retained. |
| `Submissions`, `SubmissionFiles`, `ClassroomConfig` | Classroom-backed per-assignment evidence and versioned work, or a manual teacher-entered student record in `Submissions` when there is no Classroom task. |
| `CriterionAssessments` | Criterion drafts, server-derived letters, teacher notes, and added `RubricProfileID` pin. Existing rows without a pin remain **read-only** in the web studio until deliberately reconciled. |
| `AIAssessments` | Teacher feedback draft and successful/failed AI evidence proposal metadata; no API key is stored here. |
| `AssessmentHistory` | Append-only draft, approval, import, lock and reassessment events, including approved score/feedback snapshot. |
| `ApprovedGrades` | Append-only task, student ID, rubric version, points/max, grade, teacher feedback and approval stamp. This is the **web grade register**, not a replacement for `9DAT1 Marking`. |

The spreadsheet and Drive files remain teacher-controlled. The app does not stop a direct editor of those files from manually changing cells; copy backups and audit access to the Drive library. Script-level locks and revision checks protect app-initiated writes, not arbitrary edits in Sheets. A Sheets API operation can fail partway through a multi-row change; reconcile from `AssessmentHistory`/the backup before retrying an approval if that occurs.

## Verification in this repository

```bash
npm ci
npm run check     # compile all .gs files, HTML scripts, one doGet, no duplicate legacy modules
npm test          # pure rubric scoring + mocked Sheets/Drive/Classroom/auth/UI smoke tests
npm run preview   # mock-only interface on 0.0.0.0:4173 (never stores a real key)
```

The design preview is **sample data only**. It disables entering a Gemini key/uploading real files and displays a prominent preview badge. Approval simulations in the preview never write Google data. Node tests do not constitute a live GAS deployment test; they mock Google services. Validate at narrow (390px), tablet and desktop sizes, with keyboard-only navigation and a screen reader on a real deployment.

### Live test matrix (after approval and backup)

- Allowed user loads web app; unallowed/blank user cannot call grade/import/settings APIs. Key is never echoed or written to query strings/logs. Test first with a non-student `Ping`.
- New **non-Jewellery** task (e.g. water filter) with a 10+15-mark rubric: import/enter, review, save, approve, PDF all use that task's criteria; missing marks/unknown bands cannot activate.
- Existing Jewellery task still works in its bound Sheets dialog; non-Jewellery task cannot call its 12-criterion grade API or sync through the legacy entry point. Compare legacy grade results against a backed-up workbook before deployment.
- Same student in two Classroom assignments stays in two task rows. Changed evidence creates a new version, current flag moves only after new files exist, and old approval remains intact.
- Changing an active rubric does not change an assessment already pinned to an earlier version. Approved/locked work rejects draft/AI edits; reassessment is explicit; Classroom grade is unchanged unless teacher separately uses a vetted sync path.
- Gemini may reject a model or exceed quota, and inaccessible/oversized files may be skipped; inspect the returned warning and the `AIAssessments` event, then grade manually. Test the live key only after the school's privacy policy permits it.

## Known limits / next phases

**Not yet production verified:** no live script access in this repository, no credential supplied, no real Google integration tests. This is an implementable MVP, **not a claim of production readiness**. Multi-workbook Drive enumeration can be slow for large libraries. The original sheet-bound menu uses a fixed 12-criterion Jewellery markbook; **generic approved grades are in `ApprovedGrades`** and *do not sync to Classroom*. A task-specific generic Classroom write-back (with separate teacher confirmation and verified max-point rounding), resumable batch queue, multi-model AI consensus, legacy-data reconciliation and institution-approved voice/microphone testing remain future tasks. Legacy PDF/AI/batch remain Jewellery-specific and must not be used for unrelated projects. See `docs/AUDIT.md` for baseline risks and migration phases.
