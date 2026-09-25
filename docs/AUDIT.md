# Lesson Grader (GAS) — Repository Audit & Transformation Plan

**Date:** 2026-09-25
**Audited commit:** `a468449` ("Update WebAppIntegration.gs", 2026-09-24), branch `main` = session branch base
**Scope:** Every file in the repository. All line numbers refer to the files at this commit.
**Method:** Full read of `Code.gs` (5,771 lines), `WebAppIntegration.gs` (1,277 lines) and all seven HTML files; mechanical checks (`node --check` syntax validation, `diff` of the duplicated halves of `Code.gs`, grep-based wiring traces). No production files were modified during this audit.

> **Important framing.** This audit describes **what the repository contains**, not necessarily what is deployed in Dan's live Apps Script project. As shown in §14 (Risk R1), the committed `Code.gs` **cannot parse**, so the live project must already differ from the repo. Reconciling repo ↔ live is a prerequisite for everything else (§17, Step 0).

### Unresolved inputs / assumptions (declared up front)

| # | Assumption / gap | Impact |
|---|---|---|
| A1 | The "project Markdown specification" and the "AI-agent build prompt" referenced in the tasking are **not present in this repository**. The closest artifact is `REPORT.md` (a technical evaluation + phased plan, not a product spec). | This audit treats the code + `REPORT.md` + the tasking description as the intent baseline. If a separate spec document exists, it should be added to the repo and this audit re-checked against it. |
| A2 | No access to the live Apps Script project, its bound spreadsheet, Drive folders, Classroom courses, script properties, triggers, or web-app deployment settings. | All runtime claims (e.g., "voice dictation delivery probably fails", "triggers installed") are **code-derived hypotheses**, explicitly flagged where they cannot be verified from the repo. |
| A3 | `README.md` is a single line (`# lesson-grader-gas`); there is no `appsscript.json` manifest, no `clasp` config, no tests, no CI. | The repo has no declared OAuth scopes and no deployment tooling; §15–§19 propose adding them. |
| A4 | Where the two copies of `Code.gs` disagree, Apps Script semantics mean the **later declaration wins**, so the *older* second copy (lines 3043–5771) is the "effective" logic wherever it redeclares a symbol. Statements below distinguish *intended* (first copy) vs *effective* (second copy) behaviour. | Central to the risk register and the first implementation task. |

---

## 1. Existing files and their purposes

| File | Size / lines | Purpose (verified from content) |
|---|---|---|
| `Code.gs` | 284 KB / 5,771 | The entire spreadsheet-bound backend: config, hardcoded Year 9 Jewellery rubric (C01–C12), Sheets setup/migration, roster resolution, Drive rubrics library, deterministic assessment engine, Drive file prep & PDF export, Classroom import/sync, Gemini AI service, embedded PDF-report HTML, Sheets menus, modal-dialog openers, and ~35 `api*` RPC endpoints. **The whole codebase (Sections 1–9) is present twice**: lines 1–3042 ("FULLY CORRECTED" copy + `GradeScaleService` + a mangled tail) and lines 3043–5771 (older copy). See §14 R1–R2. |
| `WebAppIntegration.gs` | 35 KB / 1,277 | Standalone web-app layer: `LessonGraderWeb` module (Drive-library traversal, workbook validation/initialisation, workbook creation, Classroom-assignment→task linking, read-only task-overview and submission-detail readers), eight `apiWeb*` endpoints, and the project's **only `doGet`** (serves `WebApp.html`). Its own header comment states it "does not make existing spreadsheet-bound grading APIs web-safe". Parses cleanly (`node --check` exit 0). |
| `WebApp.html` | 556 | The mobile-friendly ("Lesson Grader") web-app UI served by `doGet`. Views: Workbooks list/detail, Create workbook (with rubric upload), Shared rubrics, Link Classroom coursework. Explicitly displays: *"Grading still requires the Google Sheets interface; browser grading is not yet available."* Has `viewport` meta and responsive CSS. |
| `AssessmentStudio.html` | 481 | The teacher grading UI, shown as a **Sheets modal dialog** (`showModalDialog`). Two-panel desktop layout (file viewer left, 12 criterion cards right), client-side mirror of the grading maths (`deriveGradeLocal`, `recalculateOverallScore`), AI run HUD with per-file Drive preparation, feedback dock (3 FEAT-style fields), Approve/Lock/Sync/Report-PDF/Grade-Class/Dictate actions. **No `viewport` meta, no `@media` rules** → desktop-only. |
| `ClassBatchGrading.html` | 101 | Class-wide AI grading runner dialog: loads non-approved submissions (`apiGetClassSubmissions`), then sequentially calls `apiGradeSingleStudentBatch` per student with a 2.5 s pause, progress bar + log. Minified single-line script. Runs only while the dialog stays open. |
| `CourseworkPicker.html` | 288 | Classroom course → assignment picker dialog. Supports two modes via template properties `passedCourseCode` / `passedTaskName` (guarded with `typeof` checks, so missing properties don't crash): plain "save active selection" (`apiSaveClassroomSelection`) or "link to task" (`apiSaveClassroomSelectionForTask`). |
| `GradingWorkbookSetup.html` | 349 | Sheets-dialog version of workbook setup: year → course code (existing from `Rubrics Index` or new), category, task, rubric select/upload (PDF/MD/TXT read as base64), calls `apiGetGradingSetupData` / `apiSubmitGradingWorkbookSetup`, then chains into `openClassroomPickerForTask`. |
| `ManageRubrics.html` | 155 | Lists shared rubrics (`apiListSharedRubrics`), "Open in Drive" links, and a "Sync to Grading Engine" button that calls `apiRegenerateRubricSchema(rubricId)` — which regenerates the rubric file's hidden `Criteria` tab. **It does not touch the actual grading engine** (see §14 R4). |
| `VoiceDictation.html` | 146 | Dictation dialog using browser `webkitSpeechRecognition` (`lang='en-AU'`, continuous + interim). "Insert Note" posts `{type:'VOICE_TRANSCRIPT', criterionId, transcript}` to `window.opener` and closes. Contains two defects — see §9. |
| `REPORT.md` | 15 KB | Prior technical evaluation & phased plan (multi-model LLM ensemble, markdown rubric engine, roster alignment, etc.). Useful as intent; several of its factual claims about the code are outdated (e.g., it says HTML is embedded in `Code.gs` string literals — the dialogs are actually separate `.html` files; only the PDF-report template is still embedded). |
| `README.md` | 19 B | Title only. |
| `.gitignore` | 38 B | Ignores `node_modules/`, `*.py`, `*.sh`, `*.js` (with `!Code.gs` exception). Note: this will also ignore any future Jest/ESLint `.js` tooling — use `.cjs`/`.mjs` or amend the ignore rules. |

**Absent:** `appsscript.json` (manifest/scopes), any test files, CI config, clasp config, deployment docs, spec documents (A1).

---

## 2. Current user workflows (as coded)

All grading workflows start from the **bound spreadsheet's** `Assessment System` menu (`onOpen`, `Code.gs:2259`):

1. **One-time setup:** *Setup / Migrate Assessment System* → backs up the spreadsheet (Drive `makeCopy`), creates/verifies 10 system tabs, seeds `CriteriaConfig` from the hardcoded rubric, rewrites formulas on `9DAT1 Marking` (`Sheets.setupOrMigrateAssessmentSystem`, `Code.gs:421`).
2. **New task/workbook:** *Setup New Grading Workbook* (dialog) → choose year/course/category/task, select or upload a rubric (Gemini extracts PDF/MD → `Rubric` + hidden `Criteria` tabs of a new rubric spreadsheet in `Shared Rubrics`) → creates `CODE [YYYY] - Task - Grading` workbook under `Graded Assessments/Stage N/CODE/Category/` → registers the task in the `Rubrics Index` spreadsheet → optionally opens the Coursework Picker to link a Classroom assignment to the task. The same flow exists in the web app (`WebApp.html` → `apiWebCreateWorkbook`).
3. **Classroom linkage:** *Classroom → Select Course and Assignment* (dialog) → appends a `ClassroomConfig` row (`Active=true`). Task-scoped links additionally store `TaskName`, `DueDate`, `AutoImported=false`.
4. **Import submissions:** *Classroom → Import / Refresh Submissions* (manual) or hourly trigger `checkAndAutoImportDueAssignments` (auto, for rows whose `DueDate` has passed; auto path also runs AI grading per new submission). Creates/updates `Submissions` + `SubmissionFiles` rows, resolves preferred names via `ClassLists`, versions locked/approved work into reassessment rows.
5. **AI grading:** *Open Assessment Studio* → pick student → **Run AI** (per-file Drive prep with HUD, then one Gemini call with all files inline) → AI ticks observables per criterion, deterministic engine derives grades, AI generates 3-part feedback → teacher reviews/edits every checkbox and note.
6. **Batch AI grading:** *Run Class AI Grading* (dialog) → sequential per-student `apiGradeSingleStudentBatch` with progress log; dialog must stay open.
7. **Manual grading:** tick observables / Distinct-A override per criterion in the Studio; live client-side score recomputation; dictate or type per-criterion notes; edit feedback dock.
8. **Finalise:** *Save Draft* → *Approve* (writes letter grades C01–C12 + feedback into `9DAT1 Marking`, marks rows Approved) → optional *Lock* (permanent) → optional *Sync Classroom* (patches `assignedGrade`+`draftGrade` on the student submission).
9. **Reassessment:** *Administration → Create Reassessment* (prompt for a raw `SubmissionRecordID`) or automatic versioning when a locked/approved submission changes in Classroom.
10. **Reports:** *Report PDF* per student (embedded HTML → PDF data-URL download); *View Error Log*; diagnostics (*Verify Config & Sheets*, *Repair System & Re-Import*).
11. **Web app (separate entry):** deploy `doGet` → list workbooks, view metadata, list shared rubrics, create workbooks, link Classroom assignments to tasks. **No grading, no import, no AI** from the web.

---

## 3. Existing Google Sheets structures

### 3.1 Bound "grading workbook" (created/verified by `setupOrMigrateAssessmentSystem`, `Code.gs:421–477`)

| Tab | Headers (verified from schema array) | Notes |
|---|---|---|
| `Submissions` | SubmissionRecordID, ClassroomCourseID, ClassroomCourseWorkID, ClassroomSubmissionID, StudentUserID, StudentName, StudentEmail, Class, Task, SubmissionVersion, SourceType, ClassroomState, TurnedInTime, UpdateTime, Late, AttachmentSummary, AttachmentFileIDsJSON, AttachmentMetadataJSON, DriveFolderID, Status, ParentSubmissionRecordID, CurrentOfficial, LockedAt, LockedBy, ApprovedAt, ApprovedBy, ClassroomAssignedGrade, LastClassroomSyncAt, LastSyncResult, Notes (30 cols) | Status lifecycle: New → InReview → AIAssessed → TeacherReviewed → Approved → Locked (+ Superseded, SyncFailed defined but never written). `DriveFolderID` never populated. |
| `SubmissionFiles` | SubmissionRecordID, FileRecordID, SourceType, DriveFileID, FileName, MimeType, AlternateLink, ThumbnailUrl, FileSize, EligibleForAI, AIReviewStatus, AIExtractedText, Limitations, CreatedAt (14) | `AIExtractedText`/`Limitations` never written by any code path found. |
| `ClassroomConfig` | ConfigID, CourseID, CourseName, CourseSection, CourseWorkID, AssignmentTitle, MaxPoints, SavedAt, SavedBy, Active, TaskName, DueDate, AutoImported (13) | TaskName/DueDate/AutoImported appended by migration in the *first* copy only (see R2). |
| `CriteriaConfig` | CriterionID, CriterionTitle, Part, Section, MaxMarks, Outcome, Band, ThresholdType, RequiredCount, CheckboxID, CheckboxLabel, EvidenceFocus, IsMissingDistinctOverride (13) | Seeded once from the hardcoded rubric (`populateCriteriaConfig`, `Code.gs:479`). **Never read by the grading engine.** ThresholdType/RequiredCount ('All'/'TwoThirds') are unused by `deriveGradeFromChecks`. |
| `CriterionAssessments` | AssessmentID, SubmissionRecordID, CriterionID, AIProposedGrade, DerivedGrade, FinalApprovedGrade, DistinctOverrideSelected, CheckboxesJSON, AIEvidenceJSON, TeacherAdjusted, TeacherAudioTranscript, TeacherWrittenNote, Confidence, MinimumEvidenceIncomplete, Status, CreatedAt, UpdatedAt, ApprovedAt, ApprovedBy (19) | The per-criterion state store. Note: drafts already write DerivedGrade into FinalApprovedGrade (`saveCriterionDraft`, `Code.gs:1071`) — see R7. |
| `AIAssessments` | AssessmentID, SubmissionRecordID, RunType, RunTimestamp, ModelUsed, PromptVersion, InputFilesJSON, PromptSummary, ResponseJSON, ExecutionStatus, ErrorMessage, LatencySeconds, TriggeredBy, TeacherOutcome (14) | **Only `FeedbackGeneration`/`TeacherDraft` rows are ever written** (two call sites). Evidence-extraction runs are not logged → AI audit trail incomplete (R8). |
| `AssessmentHistory` | HistoryID, SubmissionRecordID, AssessmentID, Action, CriterionID, PreviousStateJSON, NewStateJSON, ActorEmail, Timestamp, Notes (10) | Written on DraftSaved / Approved / Locked / ReassessmentCreated. Import and Classroom-sync events are **not** logged. |
| `ErrorLog` | ErrorID, Timestamp, Context, ErrorMessage, StackTrace, MetadataJSON, UserEmail, ResolutionStatus (8) | API keys redacted via `Utils.sanitizeError`. |
| `ClassLists` | ClassID, OfficialName, PreferredName, SchoolEmail, Active (5) | Populated by Classroom class-list sync; `PreferredName` only ever written as `""` (manual edit expected). `ClassID` stores the Classroom **course name**. |
| `RubricProfiles` | ProfileID, ProfileName, JSONDefinition, CreatedAt (4) | **Dead schema** — created but never written or read anywhere (verified by grep). |
| `Setup` | free-form | Parsed by `GradeScaleService` (`Code.gs:28–143`) for a "GRADE SCALE" table (letter + weight/percent) and "OVERALL GRADE BAND(S)" table (letter + min [+ max]); defaults A=1.0/B=.875/C=.70/D=.575/E=.25 and bands A≥85, B≥75, C≥65, D≥50, E≥0. Layout is documented only in a code comment. |
| `9DAT1 Marking` | free-form, hardcoded name | Human markbook. Column B = student name (match key), C–N = C01–C12 letters, O/P = Part A/B totals, Q = total, R = IFS letter; AE/AF/AG (31–33) = rich-text feedback. Formulas in O–R are **rewritten** by `verifyMarkingSheetFormulas` (`Code.gs:507`) on setup and on every approval. |
| `Rubric`, `Summary` | — | Referenced in `Config` names; `Rubric` is used by the *rubric-file* readers; no code creates/uses a `Summary` tab in the workbook (legacy). |

### 3.2 `Rubrics Index` spreadsheet (in Drive, `RubricsLibrary`, `Code.gs:680`)

Tabs: `Courses` (Stage, CourseCode, Year, Active), `Categories` (CourseCode, CategoryName), `Tasks` (CourseCode, CategoryName, TaskName, RubricFileId — rubric ID immutable once set).

### 3.3 Shared rubric spreadsheets

`Rubric` tab (human-editable): Criterion, Part, Section, MaxMarks, Outcome, Band, Description. Hidden `Criteria` tab (generated by `apiRegenerateRubricSchema`, `Code.gs:2458`): same 13-column shape as `CriteriaConfig`, with generated checkbox IDs (`C01-A-1` …) and a generic MD-override text.

---

## 4. Existing Google Drive structures

Created/expected under **My Drive** (`RubricsLibrary.ensureRubricsLibraryStructure`, `Code.gs:691`):

```
Graded Assessments/
├── Settings/
│   └── Rubrics Index            (Sheets: Courses | Categories | Tasks)
├── Shared Rubrics/
│   └── <rubric name>            (Sheets: Rubric | hidden Criteria)
├── Stage 4/ ── <CourseCode>/ ── <CategoryName>/
├── Stage 5/        e.g. 9DAT1/      e.g. "Jewellery Design"/
│   └── "9DAT1 [2026] - <Task> - Grading"   (grading workbook, naming from
└── Stage 6/ ...                          RubricsLibrary.ensureGradingWorkbook:869)
```

Other Drive behaviour: spreadsheet backups `"[Backup <ts>] <name>"` via `makeCopy` on every migrate run; student attachments are **not copied** — only referenced by `DriveFileID` (Classroom-owned files); PDF export streams (`exportDriveFileAsPdf`, report PDF) return base64 data-URLs through `google.script.run`. `WebAppIntegration.gs` validates workbooks by re-walking this exact folder tree (`walk_`, `member_`) on **every** web request that opens a workbook — a per-call full Drive traversal (performance/quota risk, R12). Stage is derived from the course-code prefix (7/8→Stage 4, 9/10→Stage 5, 11/12→Stage 6, `Code.gs:2614`).

---

## 5. Existing Google Classroom integrations

All via the advanced `Classroom` service (`ClassroomService`, `Code.gs:1345–1863`):

| Integration | Function | Status |
|---|---|---|
| List teacher courses | `listTeacherCourses` (ACTIVE, `teacherId:'me'`, pageSize 50, no pagination) | Works; 50-course cap. |
| List coursework | `listCourseWork:1361` | **Two divergent versions** (R2): first copy paginates all states, sorts, tolerates null `maxPoints`; effective (second) copy fetches ≤50 PUBLISHED only and defaults `maxPoints` to 100. |
| Save selection | `saveClassroomSelection:1402` / `saveClassroomSelectionForTask:1423` | Effective copy deactivates **all** rows (incl. task rows) on a plain save, lacks column migration + task-name validation, and mis-handles `dueTime.hours===0`/`minutes===0` (falsy → 23:59). |
| Import/refresh submissions | `importOrRefreshClassroomSubmissions:1591` + `listAllStudentSubmissions:1567` (paginated; correctly omits invalid `courseWorkStates` param) | Robust pagination and roster resolution; but keys existing records by **StudentUserID only** and ignores `cfg.TaskName` → cross-task overwrite risk (R5). One `UserProfiles.get` per student (N+1 quota). Only `assignmentSubmission.attachments[].driveFile` handled — link/YouTube attachments ignored. |
| Auto-import + auto-grade on due date | `checkAndAutoImportDueAssignments:1459` + hourly trigger installer (`installAutoImportTrigger:1869`, top-level wrapper at `1865` exists only in the first copy but survives, since the second copy doesn't redeclare it) | Effective (second-copy) logic marks `AutoImported=true` **even if AI grading failed** (no retry) and silently no-ops when scheduling columns are missing. |
| Class-list sync | `syncClassListFromClassroom:1745` | Paginated students → `ClassLists`: add new (PreferredName `""`), reactivate, deactivate leavers. Never overwrites preferred names. No exception queue for unmatched emails. |
| Grade write-back | `syncAssessmentToClassroom:1825` | Requires status Approved/Locked; computes total from letter grades using **`Config.GRADE_WEIGHTS` hardcoded**, not `GradeScaleService` (inconsistency R9); patches `assignedGrade` + `draftGrade` with updateMask; records grade/time/result on the `Submissions` row. Not history-logged. |
| Web link assignment→task | `LessonGraderWeb.link` (`WebAppIntegration.gs:~465–645`) | Validates course/work, writes a TaskName-scoped `ClassroomConfig` row under `LockService`; correctly warns that triggers/manual import are separate. |

---

## 6. Existing rubric and grading logic

**Rubric source of truth (effective):** `Config.CRITERIA_DEFINITIONS` (`Code.gs:172–306`) — 12 hardcoded criteria for Year 9 Jewellery Design (C01–C12; Part A = 55 marks: C01 5, C02 10, C03 5, C04 10, C05 5, C06 10, C07 10; Part B = 45: C08 5, C09 10, C10 10, C11 10, C12 10; total 100). Each criterion: title, part, section, maxMarks, NSW outcome code (DT5-x), a "Missing/Distinct" A-override checkbox (`C0x-MD`), and five bands (A–E) × 4 observables with IDs `C0x-<Band>-<n>`.

**Deterministic scorer** `AssessmentService.deriveGradeFromChecks` (`Code.gs:922`):

1. MD override ticked → grade A, score = weight A (full marks), short-circuit.
2. Otherwise, per band: `pointPerObservable = maxMarks × weight(band) / observablesInBand`; sum over ticked observables across all bands.
3. `percentage = totalPoints / maxMarks × 100` → overall letter via `GradeScaleService.getOverallGradeLetter` (bands from `Setup` sheet or defaults; round-half-up comment; ties round up).
4. Returns `{grade, score, explanation, incompleteMinimumEvidence (zero ticks), bandBreakdown}`.

Worked example (matches REPORT.md): maxMarks 10, 3/4 Band-A ticks, weights A=1.0 → 3 × (10×1.0/4) = 7.5 → 75% → Grade B. ✔ verified by reading the code path.

**Weights/bands:** dynamic from the `Setup` sheet via `GradeScaleService` (`Code.gs:28`), used by the scorer, the marking-sheet formulas, the Studio bootstrap, and the PDF report letter. **But** Classroom sync and the PDF per-criterion marks use the hardcoded `Config.GRADE_WEIGHTS` (R9).

**Consumers of the rubric:** bootstrap data, submission detail, drafts, approval, Classroom sync, PDF report, both Gemini prompts, feedback summary — **all read `Config.CRITERIA_DEFINITIONS` directly** (14 call sites). Nothing in the grading path reads `CriteriaConfig`, the uploaded rubric files, or `RubricProfiles` (R4).

**Client-side mirror:** `AssessmentStudio.html` `deriveGradeLocal` (line ~193) re-implements the same maths for instant UI feedback using weights/bands from bootstrap — currently faithful to the server, but a second implementation that can drift.

**Markbook integration:** `Sheets.syncApprovedGradesToMarkingSheet` (`Code.gs:562`) matches the student **by name string** in column B of `9DAT1 Marking` (appends a row if not found), writes 12 letters to C–N, feedback rich text to columns 31–33, then rewrites O–R formulas for all rows.

---

## 7. Existing AI-grading behaviour

`GeminiService` (`Code.gs:1886–2166`):

- **Transport:** `UrlFetchApp` REST to `generativelanguage.googleapis.com/v1beta/models/<model>:generateContent?key=<key>` — API key is embedded in the **URL query string** (R15); key stored in Script Properties (`GEMINI_API_KEY`, set via a Sheets prompt dialog).
- **Models/fallback:** `gemini-2.5-flash` primary; fallbacks `gemini-2.0-flash`, `gemini-1.5-flash`; 3 attempts per model on 429/503 with linear backoff; `responseMimeType: application/json`, `temperature: 0.2`.
- **Inputs:** eligible files only — images (jpeg/png/webp/gif/heic) and Google Docs/Slides/Drawings (converted to PDF via `getAs`) and PDFs, ≤15 MB each, base64-inlined (`DriveService.prepareSingleFileForAi`, `Code.gs:1260`). Text is returned in one blob from `parts[0]`.
- **Two divergent evidence prompts:**
  - `runInitialAiAssessment:2030` (used by batch runner + auto-import): *"INDEPENDENT BAND TICKING (DO NOT TICK REDUNDANT LOWER BANDS)"*, asks for `suggestedBand` + `tickedCheckboxes` + `evidenceNotes`.
  - `runAiAssessmentWithPreparedFiles:1954` (used by the Studio "Run AI" button): *"TICK EVERY OBSERVABLE THE EVIDENCE SUPPORTS, ACROSS ALL BANDS"* — the opposite policy.
  Since the scorer sums points per ticked observable, **the same work can score differently depending on entry point** (R6).
- **Post-processing:** AI output is parsed; per criterion, ticks are mapped and the **deterministic engine** derives the grade (AI never computes marks — good). Drafts saved via `saveCriterionDraft` with `aiProposedGrade = suggestedBand || derived`, `confidence: 0.85` (hardcoded), status `InReview`. Unknown/invalid checkbox IDs returned by the model are stored unchecked-verified (no schema validation against the rubric).
- **Feedback:** `generateTeacherReviewedFeedback:2109` sends criterion grade summary + student's **first/full name** to Gemini with anti-jargon tone rules; returns `{whatWentWell, areasForImprovement, goalsForNextAssessment}` (FEAT-like); logs a `FeedbackGeneration` row in `AIAssessments`.
- **Human-in-the-loop:** preserved everywhere — AI only proposes; Approve/Lock/Sync are explicit teacher actions. ✔
- **Not implemented:** per-run logging of evidence extraction (R8), prompt versioning, PII stripping, prompt-injection guards, response-schema enforcement, multi-model consensus/discrepancy badges (proposed in `REPORT.md` only), streaming/cancellation.

---

## 8. Existing batch-grading behaviour

- **Entry point:** Sheets menu → `openClassAiGradingRunner` → `ClassBatchGrading.html` modal (680×420).
- **Selection:** `apiGetClassSubmissions` (`Code.gs:2925`) returns every `Submissions` row whose status is not Approved/Locked — across **all tasks/versions** in the bound workbook (no task filter, no CurrentOfficial filter).
- **Execution model:** fully **client-driven loop inside the modal**: per student, one server call `apiGradeSingleStudentBatch` = `runInitialAiAssessment` (bulk file prep + Gemini + feedback in a single execution), then `setTimeout(…, 2500)` before the next student. Progress bar + rolling log; "Stop" just halts the loop.
- **Failure semantics:** per-student failures are logged in the dialog and skipped; nothing is persisted about batch state — closing the dialog (or a browser refresh, or the Sheets session sleeping) aborts the remainder with no resume.
- **Server limits:** each per-student call must finish inside the 6-minute execution limit while base64-reading every attachment and making up to 2 Gemini calls; large folios make timeouts likely (R10). The *server-side* batch path used by the hourly trigger (`checkAndAutoImportDueAssignments`) grades inside one trigger execution for all new submissions of a row — same 6-minute exposure, and (effective copy) marks the row auto-imported even if grading failed.
- **No concurrency guard:** if the trigger fires while the dialog batch runs, both can grade the same submission (script locks exist only in `saveAssessmentDraft` and web `link_`).

---

## 9. Existing voice-dictation behaviour

- **Entry:** Studio per-criterion "Dictate" button → `google.script.run.openVoiceDictation(cid)` → `VoiceDictation.html` modal (`Code.gs:2346`).
- **Engine:** browser `SpeechRecognition`/`webkitSpeechRecognition`, `lang='en-AU'`, `continuous=true`, `interimResults=true`; graceful "unsupported browser" state.
- **Delivery:** "Insert Note" → `window.opener.postMessage({type:'VOICE_TRANSCRIPT', criterionId, transcript}, '*')`; the Studio listens on `message` and appends into the criterion's note textarea (which maps to `TeacherWrittenNote`, not `TeacherAudioTranscript` — that column stays empty in practice).
- **Defects found in code (flagged, not live-verified — A2):**
  1. `onresult` **overwrites** the textarea with only results since `event.resultIndex` → earlier phrases (and any manually typed text) are lost in continuous mode.
  2. Delivery depends on `window.opener` being the Studio frame. GAS dialogs run in sandboxed `googleusercontent.com` iframes opened as *separate* modals from the Sheets UI; `window.opener` is very likely `null`, in which case the guarded post is silently dropped and the transcript never reaches the Studio. The `TeacherAudioTranscript` schema field and history suggest an earlier server-relay design that is absent.
  3. Microphone permission inside nested GAS iframes is unreliable (no `allow="microphone"` control from here).
- **Verdict:** implemented but **probably broken end-to-end in production**; must be verified live before "preserving" it, and redesigned (server-relayed transcript or in-page recognition in the future web app) — see §18 Phase 6.

---

## 10. Existing web-app behaviour

`doGet` (`WebAppIntegration.gs:1273`) serves `WebApp.html` ("Lesson Grader"):

- **Deployment assumptions (from code, not verifiable here):** must run in a project where `Config` etc. exist; `google.script.run` executes under the deployment's execution identity. **No authentication/authorization check in `doGet` or any `apiWeb*` function** — if deployed "Anyone" + "run as me", the whole library is exposed (R13).
- **Workbooks view:** `apiWebListCandidateWorkbooks` walks `Graded Assessments/Stage 4-6/**`, validates each Sheets file has `Submissions` (SubmissionRecordID, StudentUserID, Status) and `ClassroomConfig` (all 13 headers) → "initialised" workbooks only.
- **Metadata view:** submission count per workbook.
- **Shared rubrics view:** lists rubric spreadsheets (names only; no open/edit).
- **Create workbook:** course code (stage derived server-side), task, year (20xx), rubric = none | existing | upload (PDF/MD/TXT, base64 in browser). Requires `LESSON_GRADER_WEB_CATEGORY` script property or a single unambiguous category folder. Delegates to the same `apiSubmitGradingWorkbookSetup` chain as the Sheets dialog, then initialises `Submissions`/`ClassroomConfig` in the new workbook and verifies the task tab exists.
- **Link Classroom assignment:** workbook + task tab + course + assignment → validated, locked write to `ClassroomConfig` with TaskName/DueDate/AutoImported. Honest messaging about due-date auto-import limits.
- **Built but unused:** `apiWebGetTaskAssessmentOverview` and `apiWebGetSubmissionAssessmentDetail` (read-only task roster + per-submission attachments/criterion states/history with warning arrays) have **no callers in any HTML** (grep-verified) — the read-only web review UI was never built (R3).
- **Explicitly out of scope in current UI:** grading, AI runs, imports, syncing — banner says browser grading is unavailable.
- **Mobile:** `WebApp.html` is responsive (viewport meta, fluid grid, wrapped nav) — the only mobile-friendly surface today; `AssessmentStudio.html` is not, and Sheets modal dialogs don't exist in the mobile Sheets app at all.

---

## 11. Features that are COMPLETE (present, coherent, and internally consistent in code)

| Feature | Evidence |
|---|---|
| Deterministic band-weighted scoring engine + MD override | `deriveGradeFromChecks:922`; matches REPORT.md worked example |
| Dynamic grade scale/bands from `Setup` sheet with safe defaults | `GradeScaleService:28` (defined once, not clobbered) |
| Studio criterion-checklist UI with live local recompute, notes, feedback dock | `AssessmentStudio.html` |
| Teacher-authoritative lifecycle: draft → approve → lock; reassessment versioning | `saveAssessmentDraft/approveAssessment/lockAssessment/createReassessmentFromSubmission` |
| Human-facing markbook sync incl. rich-text feedback and formula refresh | `syncApprovedGradesToMarkingSheet:562`, `verifyMarkingSheetFormulas:507` |
| Classroom import with pagination, preferred-name resolution, late/state capture, versioning of locked work | `importOrRefreshClassroomSubmissions:1591` (single-task context) |
| Classroom grade write-back (draft+assigned) gated on Approved/Locked | `syncAssessmentToClassroom:1825` |
| Gemini evidence extraction + JSON-mode + model fallback + retry/backoff | `GeminiService:1886` |
| AI-proposes / teacher-disposes separation (zero AI maths) | prompts return only checkbox IDs; engine derives |
| Per-student PDF assessment report | `HtmlTemplates.getAssessmentPdfReportHtml:2170`, `apiGenerateAssessmentPdf:2938` |
| Rubric upload → Gemini extraction → validated Rubric/Criteria sheets in Shared Rubrics | `apiCreateRubricFromDocument:2704` |
| Drive library scaffolding + Rubrics Index + workbook naming convention | `RubricsLibrary:680` |
| Class-list sync from Classroom (add/reactivate/deactivate) | `syncClassListFromClassroom:1745` |
| Error log + assessment history + API-key redaction | `Logging:388`, `Utils.sanitizeError` |
| Backup before migration | `setupOrMigrateAssessmentSystem:421` (`makeCopy`) |
| Web app: workbook discovery/validation, creation, Classroom task linking, responsive UI | `WebAppIntegration.gs` + `WebApp.html` |

## 12. Features that are PARTIAL

| Feature | What exists | What's missing |
|---|---|---|
| Multi-task workbooks | TaskName on `ClassroomConfig`, task tabs, web linking, task-scoped overview endpoints | Import ignores TaskName; submissions keyed by student only (R5); engine/markbook single-task; overview endpoints have no UI |
| Universal/dynamic rubrics | Upload, extraction, `Criteria` schema generation, `CriteriaConfig` seeding, `RubricProfiles` tab | **No wiring into the grading engine** (R4); MD-override text is generic on upload; `ThresholdType` unused |
| Web app | Setup/link/read-only server functions | No grading UI, no auth, no import/AI/sync, overview/detail endpoints unconsumed |
| Voice dictation | Dialog, recognition, transcript message | Delivery path likely broken, overwrite bug, `TeacherAudioTranscript` never persisted (A2: verify live) |
| Batch grading | Sequential runner dialog, auto-grade on import trigger | No persistence/resume, no task filter, no concurrency guard, timeout exposure, effective-copy retry bug (R2/R10) |
| AI audit trail | `AIAssessments` schema, feedback-run logging | Evidence-run logging, prompt versions, latency, input file lists, teacher outcome tracking |
| Roster alignment | `ClassLists`, preferred-name resolution on import | Manual exception queue, unmatched-email surfacing, PreferredName editing UI |
| Reassessment UX | Versioning + menu prompt by raw ID | No picker/UI; parents keep `CurrentOfficial=true` (R7c) |
| History/audit | Draft/Approve/Lock/Reassessment logged | Imports, Classroom syncs, rubric changes, trigger runs not logged |

## 13. Features that are MISSING (relative to the stated goal: complete, mobile-friendly web app)

1. **Any mobile grading surface** — the Studio is desktop-dialog-only; mobile Sheets apps can't open it.
2. **Web grading** — checklist, notes, feedback editing, approve/lock/sync via browser (server endpoints for read-only views exist; write endpoints don't).
3. **Web submission import & AI runs** (no `apiWebImport*`, no web-triggered AI).
4. **Authentication/authorization on `doGet`/`apiWeb*`** (identity check, deployment-mode guard).
5. **Rubric → engine wiring** (per-task rubric profiles consumed by scorer, Studio, AI prompts, PDF).
6. **AI run logging / prompt versioning / PII minimisation / injection guards / response schema validation.**
7. **Multi-model consensus & discrepancy badges** (REPORT.md Phase 4–5 — entirely absent from code).
8. **Resumable server-side batch queue** (job sheet + triggers + cursor).
9. **`appsscript.json` manifest, clasp tooling, tests, CI, docs** (repo hygiene).
10. **Task-aware markbook** (per-task sheets instead of the single hardcoded `9DAT1 Marking`), and Setup-sheet documentation for Dan.
11. **Exception queue UI** for unmatched roster emails.
12. **Data lifecycle**: retention/export of history, backup rotation (backups accumulate one copy per migrate run).
13. **Offline/flaky-network resilience** in any UI (no retry/queue on `google.script.run` failures except ad-hoc alerts).

---

## 14. Risks and architectural conflicts

**R1 — The committed `Code.gs` does not parse (CRITICAL, verified).** `node --check` fails at line 2960: `runEnsureRubricsLibraryStructure` ends with a stray extra brace (`}  }`), followed by a dangling fragment of `GradeScaleService`'s internals (top-level `getGradeWeights`/`getOverallGradeBands` referencing IIFE-private constants, a top-level `return`, and an unmatched `})();` at line 3037). Whatever is in Dan's live project, **the repo cannot be the deployed artifact**. Any change program must start by reconciling repo ↔ live (export the live project) and fixing this file.

**R2 — Whole-file duplication with silent override (CRITICAL, verified by diff).** Lines 3043–5771 redeclare Sections 1–9. Apps Script last-declaration-wins ⇒ the **older** copy is effective for every redeclared symbol, silently reverting fixes present in the first copy:
- `listCourseWork`: no pagination, PUBLISHED-only, `maxPoints||100`.
- `saveClassroomSelection`: deactivates **every** config row (destroys task links).
- `saveClassroomSelectionForTask`: no column migration, no task-name validation, `dueTime.hours===0` treated as absent (midnight due times become 23:59 — auto-import fires a day late... in practice: due "00:00" → stored 23:59).
- `checkAndAutoImportDueAssignments`: marks `AutoImported=true` even when grading fails (no retry); silently returns when scheduling columns missing.
- `getActiveClassroomConfig`: first-Active-row-wins instead of last-non-task row → manual import can target the wrong (task) config.
- `Sheets.setupOrMigrateAssessmentSystem`: no `ClassroomConfig` column top-up.
- `openClassroomPicker`: doesn't set `passedCourseCode/passedTaskName` (survives only because the HTML guards with `typeof`).
Symbols defined **only** in the first copy (top-level `checkAndAutoImportDueAssignments` wrapper, `apiSaveClassroomSelectionForTask`, `GradeScaleService`) survive. Net effect: an unpredictable hybrid. This is the single most dangerous defect after R1.

**R3 — Dead/duplicated surfaces.** `apiWebGetTaskAssessmentOverview`/`Detail` have no UI; `RubricProfiles` is a dead schema; `REPORT.md` describes embedded HTML strings that no longer match reality. Docs ↔ code drift will mislead future work.

**R4 — Rubric pipeline disconnected from grading (architectural conflict).** Upload → extract → `Criteria` schema stops at the rubric file. The engine, Studio, AI prompts, markbook formulas and PDF all read the **hardcoded jewellery rubric**. A workbook created for, say, "7TECHI Rube Goldberg" would import and "grade" against Year 9 Jewellery criteria. This blocks the product goal and must be fixed before any multi-task web grading.

**R5 — Import identity/scope bug (data corruption risk).** `importOrRefreshClassroomSubmissions` matches prior records by `StudentUserID` alone; with two linked tasks in one workbook, importing task B **updates/overwrites task A's row** (attachments, times) for any student not Approved/Locked. Also `Task` column = `AssignmentTitle` (fallback `'Task 2: Jewellery Design'`), never `cfg.TaskName` ⇒ web overviews filtered by task tab name will usually show zero submissions.

**R6 — Contradictory AI ticking policies** (§7) ⇒ same submission, different scores by entry point; undermines trust and calibration.

**R7 — Column-index write bugs in the lifecycle (data corruption, verified by header arithmetic).**
 a. `lockAssessment:1161` writes `now`→col 22 (**CurrentOfficial**) and user→col 23 (**LockedAt**): locking destroys the CurrentOfficial boolean and misfiles timestamps.
 b. `approveAssessment:1116` writes `now`→col 24 (**LockedBy**) and email→col 25 (**ApprovedAt**): off-by-one across three fields.
 c. Reassessment/new versions never set the parent's `CurrentOfficial=false` ⇒ multiple "current" rows.
 d. `saveCriterionDraft` pre-fills `FinalApprovedGrade` with the derived grade before any approval.
 All four are fixable with header-name→index mapping helpers + a one-time repair script (must be dry-run + backup first — §17).

**R8 — AI evidence runs unlogged.** `AIAssessments` records only feedback; there is no durable record of which model/prompt/files produced the ticks a grade was based on — weakens auditability of "preserve complete assessment history".

**R9 — Grade-maths divergence paths.** Classroom sync + PDF per-criterion marks use hardcoded `Config.GRADE_WEIGHTS`; the engine/markbook/Studio use `GradeScaleService`. If Dan edits the `Setup` grade scale, Classroom grades silently disagree with the markbook.

**R10 — GAS runtime limits.** 6-min executions vs. per-student bulk base64 + 2 Gemini calls; Studio shuttles ≤~20 MB base64 per file through `google.script.run`; `member_()` walks the whole Drive library per web call; `getSheetDataAsObjects` full-sheet reads repeated within one request; N+1 `UserProfiles.get`. Class-scale (30 students × multi-MB folios) will hit timeouts/quotas.

**R11 — Bound-spreadsheet vs library-of-workbooks conflict.** Every grading API uses `SpreadsheetApp.getActiveSpreadsheet()` ⇒ works only in the one bound workbook (or its container when the bound script is deployed as a web app). The Drive library/web app imply *many* workbooks. Until every engine API takes an explicit `spreadsheetId`, the web app can never grade more than the bound workbook (`WebAppIntegration.gs`'s own header admits this).

**R12 — Concurrency.** Script locks only in `saveAssessmentDraft` and web `link_`. Trigger + dialog + web can race on `ClassroomConfig`/`Submissions` writes.

**R13 — Security/privacy.** No auth check in `doGet`; execution-identity exposure if mis-deployed; Gemini API key in URL query (server logs/error strings; partially redacted); student names + work sent to Gemini without PII minimisation; `postMessage('*')`; whether sending student work to the Gemini API complies with Dan's school/sector policy is **unknown to me and must be confirmed** (A2).

**R14 — Hardcoding ceiling.** `9DAT1 Marking`, `9DAT1`, `Task 2: Jewellery Design`, fixed columns C–N/O–R for exactly 12 criteria, PDF title/`/55`/`/45`/filename — every "general" feature eventually collides with these.

**R15 — Fragile markbook overwrite.** `verifyMarkingSheetFormulas` rewrites O–R on every approval; any manual customisation Dan made there is silently destroyed; student matching by display name is collision-prone (preferred names make this worse).

**R16 — Repo/process risk.** No manifest ⇒ scopes are whatever the live project accumulated; no tests ⇒ every refactor is unverifiable; `.gitignore` blocks `.js` tooling; single-commit history ⇒ no provenance for the duplication.

---

## 15. Recommended target architecture

Constraints honoured: Google Classroom/Drive/Sheets remain the system of record; Dan remains final authority; Apps Script (free, serverless, NSW-school-friendly); incremental, reversible steps; mobile-first web UX.

```
                      ┌────────────────────────────────────────────────┐
                      │        Apps Script project (standalone,        │
                      │  appsscript.json + clasp + CI syntax/tests)    │
                      │                                                │
  Mobile/desktop ───► │  doGet → WebApp.html (single mobile-first SPA) │
  browser (Dan only)  │    views: Tasks ▸ Roster ▸ Student ▸ Review    │
                      │    all calls carry {workbookId, taskName}      │
                      │                                                │
                      │  Api layer (thin, validated, auth-checked)     │
                      │   ├─ WebSetupApi   (workbooks, rubrics, links) │
                      │   ├─ GradingApi    (bootstrap/detail/draft/    │
                      │   │                 approve/lock/reassess/sync)│
                      │   ├─ AiApi         (run/feedback/status)       │
                      │   └─ BatchApi      (enqueue/cursor/results)    │
                      │                                                │
                      │  Domain services (workbook-scoped, no          │
                      │  getActiveSpreadsheet):                        │
                      │   RubricEngine (RubricProfiles JSON per task)  │
                      │   ScoringEngine (pure functions, unit-tested)  │
                      │   AssessmentStore (Sheets rows, header-mapped) │
                      │   ClassroomGateway (import/sync, task-scoped)  │
                      │   RosterService (ClassLists + exception queue) │
                      │   AiGateway (provider-agnostic; Gemini now,    │
                      │     ensemble later; run logging; schema check) │
                      │   BatchQueue (BatchJobs sheet + time trigger,  │
                      │     resumable cursor, per-item error capture)  │
                      │   AuditLog (AssessmentHistory + AIAssessments) │
                      └───────────────┬────────────────────────────────┘
                                      │ source of truth (unchanged)
              Google Sheets workbooks • Drive "Graded Assessments" library • Classroom
```

Key decisions (each reversible, each justified by a finding above):

1. **Standalone script project + explicit `workbookId` on every API** (fixes R11); keep a tiny `onOpen`-style helper only if Dan still wants Sheets menus, implemented via a *bound* companion script that calls the standalone web app — or, simpler for Phase 1–3, stay bound but thread `spreadsheetId` through all domain calls so the code is location-independent. Recommendation: **convert to standalone** once APIs are parameterised; deploy `doGet` restricted to Dan's account ("Execute as: me", "Who has access: only myself") and add an email allowlist check server-side regardless (fixes R13).
2. **One rubric profile per task** stored as JSON in `RubricProfiles` (canonical), with `CriteriaConfig` as a derived flat view for humans/AI prompt serialisation. Scoring engine, Studio/web UI, AI prompts, PDF and markbook formulas all read the profile (fixes R4, R14). Hardcoded jewellery rubric becomes a *seeded profile*, not code.
3. **Single evidence-extraction prompt policy** (choose "tick every supported observable across bands" — it matches the per-observable scorer's semantics) behind a `promptVersion` constant; every run logged to `AIAssessments` (fixes R6, R8).
4. **Server-side batch queue** (`BatchJobs` sheet + 5-min time trigger + cursor) replacing dialog-driven loops; Studio/web only enqueue and poll (fixes R10 partially, resumability).
5. **Header-name→index mapping helper** for all sheet writes; column numbers never appear in business logic (fixes R7 class of bugs permanently).
6. **Provider-agnostic `AiGateway`** so REPORT.md's multi-model consensus (Gemini/Qwen/DeepSeek) can be added later *without* touching the scoring engine — but deliberately **deferred** until the web MVP is stable; nothing in the repo implements or requires it today.
7. **Voice dictation redesign:** in-page `SpeechRecognition` inside the web app (single frame — no `window.opener` hop), transcript saved through the normal draft API into `TeacherAudioTranscript`; keep the GAS dialog version for the legacy Sheets UI until it is retired (fixes §9 defects, pending live verification A2).
8. **HTML served from files** (`createTemplateFromFile`) — already true for dialogs; the web app stays a single responsive file with progressive enhancement; Studio logic (client scorer) becomes a shared, unit-tested JS module used by the web app, with the server remaining authoritative on every save.

## 16. Proposed data model

Unchanged sources of truth: Classroom (roster, submissions, grades), Drive (files, library tree), Sheets (all state). Changes are additive unless marked *repair*.

**Per-workbook tabs (evolved):**

- `Submissions` — keep 30 columns; **add** `TaskName` (populated from `ClassroomConfig.TaskName` at import); *repair* semantics: uniqueness of "current version" per (`StudentUserID`, `ClassroomCourseWorkID`); `CurrentOfficial` lifecycle enforced (exactly one true per student+task; parents demoted on versioning). Index helper sheet optional.
- `SubmissionFiles` — unchanged; start populating `AIReviewStatus`/`Limitations` from `prepareSingleFileForAi` results (fields already exist).
- `ClassroomConfig` — unchanged; **enforce** one `Active=true` per `TaskName` scope; `(CourseID, CourseWorkID, TaskName)` unique.
- `CriterionAssessments` — unchanged columns; *semantics fix*: `FinalApprovedGrade` written **only** on approval; add `PromptVersion` + `AiRunIDsJSON` (or rely on AIAssessments join via AssessmentID).
- `AIAssessments` — unchanged columns; **all** runs logged: `RunType ∈ {EvidenceExtraction, FeedbackGeneration, RubricExtraction}`, `PromptVersion`, `InputFilesJSON`, `LatencySeconds`, `ExecutionStatus`, `TeacherOutcome ∈ {Accepted, Adjusted, Rejected}` (backfilled at approve time by diffing ticks).
- `AssessmentHistory` — unchanged; **add** actions: `Imported`, `ClassroomSynced`, `BatchQueued/BatchCompleted`, `RubricProfileChanged`, `DataRepair` (with actor + notes) so history stays complete.
- `RubricProfiles` — **activated**: `ProfileID, TaskName, ProfileName, SourceRubricFileId, JSONDefinition, SchemaVersion, PromptPolicy, Active, CreatedAt, CreatedBy`. `JSONDefinition` follows the canonical schema already drafted in `REPORT.md` §5 (criterionId, title, maxMarks, masteryOverride, bands{A..E:[{id, observableText}]}) extended with `part`, `section`, `outcome` (needed by PDF/markbook/Classroom sync).
- `CriteriaConfig` — becomes a **derived** human-readable flatten of the active profile (regenerated, never hand-edited; header notes this).
- `ClassLists` — unchanged; **add** `MatchStatus` (`matched|exception`) or a separate `RosterExceptions` tab (`SchoolEmail, ClassroomName, CourseID, FirstSeen, Resolved, ResolvedTo, Notes`) fed by import when `resolveStudentIdentity.matched === false`.
- **New** `BatchJobs`: `JobID, Scope(workbookId/taskName/filter), State(Queued|Running|Paused|Done|Failed), CursorSubmissionID, TotalItems, ProcessedItems, FailedItemsJSON, StartedAt, UpdatedAt, TriggeredBy`.
- **New** `Settings` (or documented `Setup`): grade scale + bands (existing), plus `PromptPolicy`, `AiProviderConfig`, feature flags. Keep `Setup` backward-compatible; document its layout in the repo (currently only a code comment).
- Markbook: replace single `9DAT1 Marking` with **per-task marking tabs** (`<TaskName> Marking`) generated from the active profile (columns derived from criteria count, not fixed C–N); feedback columns positioned by header lookup, not constants. Migration keeps `9DAT1 Marking` untouched and readable (history preservation).

**IDs & integrity:** keep UUID primary keys; all writes under `LockService` per workbook; all timestamps ISO via `Utils.formatDate`; every mutation appends to `AssessmentHistory`.

## 17. Migration strategy

Guiding rules: backup before every structural change (existing `makeCopy` pattern, extended); additive columns only; no destructive edits to `AssessmentHistory`/`AIAssessments`; dry-run + report for data repairs; feature-flag new paths via Script Properties; keep the Sheets-dialog UX working until the web app reaches parity (parallel run).

- **Step 0 — Reconcile & stabilise (no live changes).** Export Dan's live Apps Script project (or `clasp pull`) and diff against the repo; decide canonical content (evidence says first copy = intended "FULLY CORRECTED" version). Fix R1/R2 **in the repo first** (remove the stale second copy + dangling tail; keep first-copy improvements; re-add anything that only exists live). Add `appsscript.json` with least-privilege scopes, `.clasp.json`, and a CI syntax check. Only push to the live project after Dan confirms and a spreadsheet backup exists.
- **Step 1 — Data repair (guarded).** One-time `repairLifecycleColumns()` script: read `Submissions`/`CriterionAssessments` by header name; detect R7 damage (timestamps in boolean/letter columns, duplicated `CurrentOfficial`); output a dry-run report sheet; apply fixes under lock; log `DataRepair` history rows. Never rewrites `AssessmentHistory`.
- **Step 2 — Schema additions.** Add `TaskName` to `Submissions`, create `BatchJobs`, `RosterExceptions`, activate `RubricProfiles` (all additive; `setupOrMigrateAssessmentSystem` extended, idempotent).
- **Step 3 — Rubric profile seeding.** Generate the jewellery profile from `Config.CRITERIA_DEFINITIONS` into `RubricProfiles` (Active) + regenerate `CriteriaConfig` from it; run golden tests (§19) proving identical grades for identical tick sets before switching the engine to profile reads; keep `Config` as emergency fallback behind a flag.
- **Step 4 — Task-scoped import.** Import writes `TaskName`, matches prior rows by (`StudentUserID`,`CourseWorkID`); backfill `TaskName` on existing rows from `ClassroomConfig` where unambiguous, else mark for manual review (report, don't guess).
- **Step 5 — Web app growth.** Wire the already-built overview/detail endpoints into `WebApp.html`; add grading write endpoints (reusing domain services with `workbookId`); add auth check; keep dialogs functional throughout.
- **Rollback:** every step is independently deployable and revertible via git + prior script version in the Apps Script version history; data steps have backups + dry-run reports. No step deletes a tab or column.

## 18. Implementation phases (small, testable increments)

| Phase | Content | Exit criteria |
|---|---|---|
| **0 — Stabilise repo** | Fix R1 syntax error + R2 duplication (keep corrected copy); add manifest, clasp, eslint, CI `node --check`-style gate; README/docs refresh; reconcile with live export | Repo parses; CI green; live project matches repo after Dan-approved push; zero behaviour change beyond the intended corrected copy |
| **1 — Data integrity** | R7 column-write fixes via header-mapped writes; CurrentOfficial lifecycle; locks on write paths; one-time repair script (dry-run first); history for imports/syncs | Unit tests for header mapper + lifecycle; repair report reviewed by Dan; no regression in Studio flows |
| **2 — Dynamic rubric engine** | RubricProfiles seeding + schema; engine/AI/PDF/markbook read profiles; golden-parity tests vs hardcoded behaviour; per-task marking tabs | Jewellery task grades bit-identical; a second dummy task grades against its own uploaded rubric end-to-end |
| **3 — Mobile web grading MVP** | Auth-checked `doGet`; task→roster→student navigation consuming existing overview/detail endpoints; grading checklist + notes + feedback with server-authoritative scoring; approve/lock/sync from web; responsive pass on real devices | Dan can grade one real submission fully on a phone; Sheets dialog remains as fallback |
| **4 — AI on web + auditability** | Unified evidence prompt + promptVersion; log every AI run; server-side file prep (no base64 via browser); web "Run AI" + "Generate feedback"; TeacherOutcome backfill at approve | AI runs fully auditable in `AIAssessments`; web AI run within execution limits for a typical folio |
| **5 — Resumable batch** | `BatchJobs` queue + time-trigger worker + cursor; web/dialog UI shows progress; per-student failure isolation; concurrency guard vs manual grading | 30-student batch survives dialog close/refresh; failures retryable; no duplicate grades |
| **6 — Roster, voice, reassessment UX** | Exception queue UI; PreferredName editing; in-page dictation for web app (transcript → `TeacherAudioTranscript`); reassessment picker (no raw UUID prompts) | Unmatched students visible & resolvable; dictation verified working on Dan's actual devices/browsers |
| **7 — Hardening & optional ensemble** | PII minimisation, key via header, quota dashboards, backup rotation, retention policy; **only then** evaluate REPORT.md's multi-model consensus behind `AiGateway` (Qwen/DeepSeek discrepancy badges) | Security checklist signed off; ensemble strictly additive & flag-gated if pursued |

Phases 0–1 are prerequisites for everything; 2–3 deliver the headline product goal (mobile-friendly web app grading real tasks); 4–7 are quality/scale. Each phase is independently shippable and reversible.

## 19. Testing strategy

1. **Golden-behaviour capture first.** Before any refactor, extract the current scorer (`deriveGradeFromChecks` + `GradeScaleService` parsing) into pure functions and snapshot outputs for a fixture matrix (all-bands-empty, MD override, single ticks per band, mixed bands, custom Setup weights/bands, malformed Setup tables). These snapshots are the contract for Phase 2 ("never silently alter approved assessments").
2. **Unit tests (Node/Jest + GAS mocks).** `gas-local` or hand-rolled mocks for `SpreadsheetApp/DriveApp/Classroom/UrlFetchApp/PropertiesService/LockService/Session/Utilities`. Targets: scoring engine, grade-scale parser, header-mapper, identity resolution, due-date normalisation (incl. `hours===0` regression), rubric JSON validation, Classroom-config state machine, batch cursor logic. Note `.gitignore` blocks `*.js` — use `.cjs`/`.mjs` or amend.
3. **Static gates in CI.** Syntax check every `.gs` (concatenate → parse), ESLint, manifest scope diff alert, duplicate-top-level-symbol detector (would have caught R2), dead-endpoint report (would have caught R3).
4. **Fixture/integration tests against a throwaway workspace.** A test Drive folder + test Classroom course + test workbook; scripted flows: create workbook → upload rubric → link assignment → import → AI (recorded/mocked UrlFetch responses) → draft → approve → sync; assert exact sheet cell states by header lookup. Re-run after each phase.
5. **Data-repair rehearsal.** Run Step-1 repair in dry-run against a *copy* of the real workbook; Dan reviews the diff report before apply.
6. **UI tests.** jsdom tests for the client-side scorer parity with server; manual mobile device matrix (iOS Safari, Android Chrome, tablet) per phase; accessibility spot-checks (existing `aria-live`, focus styles are a good base).
7. **Live-verification checklist for unverifiable-from-repo items (A2).** Voice dictation delivery, microphone permission, current deployment settings ("run as"/"who has access"), installed triggers, actual `Setup` sheet contents, actual `9DAT1 Marking` layout, whether the live script matches the repo. Each item gets an explicit ✔/✘ before the phase that depends on it.
8. **Regression policy.** Every bug in this audit (R1–R16) becomes a named test before its fix lands.

## 20. The first small implementation task (proposed — awaiting go-ahead)

**Task: "De-duplicate and repair `Code.gs`; add manifest + CI syntax gate."** (Phase 0, step 1 — repo-only; touches no live system.)

1. Remove the stale second copy (`Code.gs` lines 3043–5771) and the mangled seam (stray `}` at line 2960 + dangling `GradeScaleService` fragment 2962–3037), keeping the first copy — the self-described "FULLY CORRECTED" version whose improvements (paginated `listCourseWork`, task-safe `saveClassroomSelection`, retry-safe auto-import, column-migrating setup, task-aware `getActiveClassroomConfig`) are currently being silently overridden.
2. Re-verify: `node --check` passes; a duplicate-symbol scan reports zero collisions; every `api*`/menu/trigger handler referenced by the HTML files and `installAutoImportTrigger` still exists exactly once (diff-driven checklist, incl. `apiSaveClassroomSelectionForTask` and top-level `checkAndAutoImportDueAssignments`).
3. Add `appsscript.json` declaring the scopes actually used (Sheets, Drive, Classroom, UrlFetch, ScriptApp, PropertiesService, `script.external_request`), `.clasp.json` template, and a minimal CI workflow (syntax + duplicate-symbol checks).
4. Produce a short reconciliation note for Dan: because the committed file cannot parse (R1), his live project necessarily differs — before deploying the repaired file, export the live project and diff; the repaired file must then be pushed only after a spreadsheet backup.

**Why this first:** it is small, mechanical, fully reversible (git), verifiable offline, changes no data, and unblocks every later phase — no feature work should be layered on a file that doesn't parse and whose effective behaviour is an accidental hybrid of two versions. **Estimated blast radius if skipped:** every subsequent fix risks being silently overridden by the stale copy (R2) or being unverifiable (R16).

---

### Appendix A — Verification commands used in this audit

```bash
node --check Code.gs               # → SyntaxError at line 2960 (R1)
node --check WebAppIntegration.gs  # → OK
diff <(sed -n '145,3042p' Code.gs) <(sed -n '3043,5771p' Code.gs)   # R2: 17 hunks
grep -n 'RubricProfiles' *.gs      # schema only — never read/written (R4)
grep -rn 'apiWebGetTaskAssessmentOverview\|apiWebGetSubmissionAssessmentDetail' *.html  # no callers (R3)
grep -n 'CRITERIA_DEFINITIONS' Code.gs   # 14 grading-path consumers (R4)
```

### Appendix B — Effective-vs-intended divergence map (R2 detail)

| Symbol | First copy (intended, lines) | Second copy (effective, lines) | Practical delta |
|---|---|---|---|
| `listCourseWork` | 1361 | ~4259 | pagination/sorting/null-safety lost; ≤50 PUBLISHED |
| `saveClassroomSelection` | 1402 | ~4300 | all rows deactivated (task links wiped) |
| `saveClassroomSelectionForTask` | 1423 | ~4321 | no column migration/validation; midnight-due bug |
| `checkAndAutoImportDueAssignments` (method) | 1459 | ~4357 | marks AutoImported even on grading failure; silent skip |
| `ensureClassroomConfigColumns_` | 1540 | absent | helper gone from effective module |
| `getActiveClassroomConfig` | 1565 | ~4400 | first-Active wins; may pick task rows |
| `Sheets.setupOrMigrateAssessmentSystem` | 421 | ~3319 | no ClassroomConfig column top-up |
| `openClassroomPicker` | 2324 | ~5138 | template props unset (guarded in HTML) |
| `apiSaveClassroomSelectionForTask` (global) | 2907 | absent | survives (defined once) |
| top-level `checkAndAutoImportDueAssignments` (trigger target) | 1865 | absent | survives (defined once) |
| `GradeScaleService` | 28 | absent | survives (defined once) |
