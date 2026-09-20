# Comprehensive Technical Evaluation & Phased Implementation Plan

## 1. Executive Assessment
### Current State
The `lesson-grader-gas` repository contains a monolithic Google Apps Script (GAS) application intended to grade Year 9 Jewellery Design assignments. It is heavily hardcoded to specific criteria (C01–C12) and uses Google Sheets as a primary backend and Google Classroom for synchronization. It features embedded UI templates (HTML/JS/CSS) served via GAS `HtmlService`.

### Operational Strengths
*   **GAS Native:** Tightly integrated with Google Workspace (Classroom, Sheets, Drive).
*   **Serverless Foundations:** Zero local inference footprint; uses `UrlFetchApp` for remote AI APIs.
*   **Structured Storage:** Uses Google Sheets logically as a relational datastore (`Submissions`, `SubmissionFiles`, `CriterionAssessments`, `AssessmentHistory`).
*   **Human-in-the-loop:** UI heavily prioritizes teacher override, approval, and final grading before Classroom sync.

### Technical Debt & Key Architectural Bottlenecks
*   **Hardcoded Models & Prompts:** The rubric criteria, marks, outcomes, and grading bands (A-E) are hardcoded into `Config.CRITERIA_DEFINITIONS` in `Code.gs`.
*   **Monolithic Model:** A single AI provider (Gemini via Google AI Studio) is used for all evaluation, which restricts specific auditing capabilities (e.g., Qwen for specific vision tasks, DeepSeek for reasoning).
*   **Rigid Roster Resolution:** `importOrRefreshClassroomSubmissions` pulls blindly from Google Classroom without normalized preferred-name mapping or manual exception queues.
*   **Fragile UI Code:** HTML files are bundled in large string literals in `HtmlTemplates` within `Code.gs`. They are hard to maintain, test, and scale.

### Primary Architectural Shift Needed
Transition from a **hardcoded, single-project grader** to a **universal, data-driven multi-model platform**. This means moving criteria definitions into a dynamic Markdown parser, utilizing a tiered cloud LLM ensemble (Gemini + Qwen + DeepSeek), and solidifying the deterministic grading engine inside GAS.

## 2. Codebase Inventory & Technical Audit

| File Path | Primary Role | Key Exported Symbols | External Integrations | Technical Risks |
| :--- | :--- | :--- | :--- | :--- |
| `Code.gs` | Monolithic backend script | `Config`, `GradeScaleService`, `Sheets`, `AssessmentService`, `GeminiService`, `DriveService`, `ClassroomService`, `HtmlTemplates` | Google Sheets API, Google Classroom API, Google Drive API, Gemini API | 1200+ lines monolithic script, hardcoded criteria (`CRITERIA_DEFINITIONS`), embedded HTML strings bypassing separate HTML files. |
| `AssessmentStudio.html` | UI template | N/A | N/A | Client-side grading logic duplicate. |
| `CourseworkPicker.html` | UI template | N/A | N/A | Exists as separate file. |
| `VoiceDictation.html` | UI template | N/A | N/A | Exists as separate file. |

**Detailed Audit of Code.gs:**
*   **Data Models:** Represented dynamically in Google Sheets tabs (e.g., `SHEET_SUBMISSIONS`, `SHEET_CRITERIA_CONFIG`).
*   **Config Constants:** Heavily hardcoded in `Config`, specifically `CRITERIA_DEFINITIONS` which spans ~300 lines defining Year 9 Jewellery Design.
*   **Spreadsheet Interaction Patterns:** `Sheets.getSheetDataAsObjects` is used heavily, reading entire data ranges into memory, which scales poorly for large datasets or execution limits.
*   **AI Integration:** `GeminiService` hardcodes Gemini prompts (`getSerializedRubricPrompt`) and expects rigid JSON returns.

## 3. Gap & Risk Analysis
*   **Universal Rubric Engine:** *Absent*. Rubrics are currently hardcoded JS objects.
*   **Specialist Multi-Tier Cloud LLM Ensemble:** *Absent*. Only Gemini 2.5/1.5 Flash is configured. No discrepancy engine exists.
*   **Deterministic Grading:** *Implemented*. `AssessmentService.deriveGradeFromChecks` calculates deterministically.
*   **Classroom Roster Alignment:** *Absent*. Relies strictly on Classroom User Profiles (`Classroom.UserProfiles.get`), ignoring spreadsheet-based master rosters or preferred names.
*   **Apps Script Runtime Limits:** *Risky*. Batch processing (`importOrRefreshClassroomSubmissions` and `apiGradeSingleStudentBatch`) risks 6-minute timeout on large classes or large PDF ingestions.
*   **Prompt Injection / Privacy:** *Risky*. Raw student text is sent to Gemini without explicit PII stripping or prompt injection guards.

## 4. Target Architecture & Specialist Cloud LLM Tiering

### ASCII Architectural Pipeline Diagram
```
[Google Classroom / Drive Ingestion]
         |
         v
[Stage 1: Multimodal Extraction] ----> (Gemini 2.5 Flash: Bulk Context & General Observables)
         |
         v
[Stage 2: Visual & Logical Audit] ---> (Qwen2.5-VL: Physical Fabrication / Sketches)
         |                        ---> (DeepSeek-R1: Citation / Logic Verification)
         v
[Stage 3: Discrepancy Engine] -------> (GAS: Compares AI Outputs. Consensus = High Conf, Divergence = Review Badge)
         |
         v
[Stage 4: Deterministic GAS Scoring]-> (GAS: Weights, Overrides, Mark Allocation. ZERO AI Math)
         |
         v
[Stage 5: Teacher Review Studio] ----> (GAS HtmlService / Vue or Vanilla JS: Final human approval & sync)
```

### External Cloud Models Specification Table
| Role | Model Name | Context/Vision Profile | Hosting/Endpoint | Quota/Cost | GAS Calling Method |
| :--- | :--- | :--- | :--- | :--- | :--- |
| Primary Scanner | Gemini 2.5 Flash | 1M+ tokens, Multimodal | Google AI Studio | Free Tier / Low Cost | `UrlFetchApp` (REST) |
| Visual Auditor | Qwen2.5-VL (7B/8B) | Image/Vision heavily tuned | OpenRouter / HF Spaces | Fractional cents | `UrlFetchApp` (REST) |
| Logical Verifier | DeepSeek-R1 Distill | CoT, Logical Reasoning | OpenRouter / API | Fractional cents | `UrlFetchApp` (REST) |

*Explicit Boundary:* LLMs **ONLY** return arrays of checked observable IDs and citations (JSON Schema). GAS executes all logic to convert IDs -> Scores -> Grades.

## 5. Universal Rubric Engine & Calibration Pipeline

### Canonical JSON Schema
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "type": "object",
  "properties": {
    "profileId": { "type": "string" },
    "criteria": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "criterionId": { "type": "string" },
          "title": { "type": "string" },
          "maxMarks": { "type": "number" },
          "masteryOverride": { "type": "object", "properties": { "id": {"type": "string"}, "text": {"type": "string"} } },
          "bands": {
            "type": "object",
            "additionalProperties": {
              "type": "array",
              "items": {
                "type": "object",
                "properties": {
                  "id": { "type": "string" },
                  "observableText": { "type": "string" }
                }
              }
            }
          }
        }
      }
    }
  }
}
```

### Markdown Ingestion Pipeline
1.  **Raw Storage:** Teacher pastes Markdown rubric into a UI text area.
2.  **Deterministic Pre-parser:** Regex/AST parsing in GAS extracts Table rows/bullets into structured Criteria, Bands, and Observables.
3.  **LLM Extraction (Optional):** If markdown is unstructured, Gemini structures it against the JSON schema.
4.  **Multi-Judge Stress-Test / Calibration:** DeepSeek analyzes the extracted observables for semantic overlap and ambiguity (e.g., "Is Band B observable 2 logically distinct from Band C observable 1?").
5.  **Teacher Review & Commit:** Saves to `CriteriaConfig` Google Sheet.

### Scoring Algorithm
```javascript
function calculateScore(criterionId, checkedObservableIds, rubricProfile, gradeWeights) {
    const criterion = rubricProfile.criteria.find(c => c.criterionId === criterionId);

    // Distinct Mastery Override always yields full marks
    if (checkedObservableIds.includes(criterion.masteryOverride.id)) {
        return criterion.maxMarks * gradeWeights['A'];
    }

    const bandCalcs = {};
    const bandNames = ['A', 'B', 'C', 'D', 'E'];

    // Calculate fraction of observables met per band
    for (let bName of bandNames) {
        const items = criterion.bands[bName] || [];
        let ticked = 0;
        for (let item of items) {
            if (checkedObservableIds.includes(item.id)) ticked++;
        }
        bandCalcs[bName] = items.length > 0 ? (ticked / items.length) : 0;
    }

    // Build deterministic score via additive weighted fractions
    // Base score is 'E'
    let score = gradeWeights['E'];
    score += bandCalcs['D'] * (gradeWeights['D'] - gradeWeights['E']);
    score += bandCalcs['C'] * (gradeWeights['C'] - gradeWeights['D']);
    score += bandCalcs['B'] * (gradeWeights['B'] - gradeWeights['C']);
    score += bandCalcs['A'] * (gradeWeights['A'] - gradeWeights['B']);

    // Cap at 'A'
    if (score > gradeWeights['A']) score = gradeWeights['A'];

    return criterion.maxMarks * score;
}
```

## 6. Class, Roster & Preferred Name Alignment Design

### Data Schema (ClassLists Tab)
| ClassID | OfficialName | PreferredName | SchoolEmail | Active |
| :--- | :--- | :--- | :--- | :--- |
| 9DAT1 | John Smith | Jack | john.smith@student.edu.au | TRUE |

### Identity Resolution Algorithm
1.  On Google Classroom Sync, extract Google `userId` and `emailAddress`.
2.  Normalise email (lowercase, trim).
3.  Lookup normalized email in `ClassLists` sheet.
4.  If match -> use `PreferredName` for UI, Feedback Gen, and Exports.
5.  If NO match -> Flag in "Manual Exception Queue" (UI state) for teacher to map manually. Do NOT fuzzy match to avoid FERPA/PII cross-contamination.

## 7. Evidence Verification, Discrepancy & Feedback Schema

### Evidence Verification Schema
```json
{
  "criterionId": "C01",
  "verdicts": [
    {
      "observableId": "C01-B-1",
      "model": "gemini-2.5-flash",
      "confidence": 0.92,
      "citation": "Slide 4, visual sketch",
      "status": "checked"
    },
    {
      "observableId": "C01-B-1",
      "model": "qwen2.5-vl",
      "confidence": 0.88,
      "citation": "Slide 4, hand drawn lines detected",
      "status": "checked"
    }
  ],
  "consensusStatus": "HIGH_CONFIDENCE"
}
```

### Discrepancy Matrix
*   **Gemini (Checked) + Qwen (Checked) = High Confidence.** Checkbox ticked in UI.
*   **Gemini (Checked) + Qwen (Unchecked) = Divergence.** Checkbox unticked in UI, flagged with Amber "Review Needed" badge.
*   **Gemini (Unchecked) + Qwen (Unchecked) = High Confidence Negative.** Checkbox unticked.

### FEAT Feedback System Prompt Arch
```text
System: Act as an Australian Secondary TAS teacher.
Review the following consensus observables.
Using the FEAT framework:
1. Praise (Strengths based on A/B observables achieved).
2. Actionable Remediation (Based on missing C/D observables).
3. Next Milestone (Project specific next step).
Do not use generic corporate AI jargon.
```

## 8. Data Storage, Sheet Migration & Operations

### Proposed Google Sheet Tab Structures
*   `RubricProfiles` (Metadata, ID, JSON definition string)
*   `CriteriaConfig` (Migrated to relate to RubricProfileId)
*   `ClassLists` (Roster mapping)

### Safe Migration Strategy (`setupOrMigrateAssessmentSystem`)
*   Append new columns to existing sheets (e.g., `RubricProfileID` to `CriteriaConfig`).
*   Backfill existing hardcoded criteria into the new `RubricProfiles` format.
*   Implement `LockService` strictly around sheet writes to prevent concurrent execution overwrites during batch processing.

## 9. Phased Implementation Roadmap

*   **Phase 1: Core Stabilization & Schema Migration Foundations.** Clean up monolithic `Code.gs`. Separate HTML files. Build `ClassLists` and `RubricProfiles` sheet structures. *Acceptance: Clean build, zero regression on existing marking.*
*   **Phase 2: Roster Alignment & Preferred Name Integration.** Implement identity resolution. Update UI to use preferred names. *Acceptance: Teacher sees preferred names in studio; unknown emails go to queue.*
*   **Phase 3: Markdown Rubric Ingestion & Dynamic Scoring Engine.** Build Markdown ingestion UI. Port `deriveGradeFromChecks` to use dynamic profiles. *Acceptance: Teacher can paste markdown and test deterministic score.*
*   **Phase 4: Specialist Cloud LLM Integration.** Implement OpenRouter REST calls for Qwen and DeepSeek. *Acceptance: UrlFetchApp calls succeed and return schema-valid JSON.*
*   **Phase 5: Evidence Consensus & Assessment Studio UI Enhancements.** Implement discrepancy logic. Add "Amber Review" badges to UI. *Acceptance: Conflicting AI verdicts require explicit teacher click.*
*   **Phase 6: Rubric Calibration Council & Enhanced FEAT Feedback Engine.** DeepSeek stress-tests rubrics. FEAT prompt engineering. *Acceptance: Feedback reads naturally; rubric parser flags overlapping criteria.*

## 10. Test Strategy & Quality Assurance
*   **Unit Tests:** Pure JS functions in GAS (e.g., `calculateScore`, `resolveIdentity`, `parseMarkdownRubric`) tested locally using a mock GAS environment (e.g., Jest with GAS mocks).
*   **Fixture Tests:** Supply malformed JSON, ambiguous markdown, and edge-case arrays to the discrepancy engine.
*   **Edge Cases:** Handle 6-min GAS timeouts by chunking API calls. Handle missing `ClassLists` records.

## 11. Strategic Decisions & Exclusions
*   **Recommended Decisions:** Shift HTML templates to actual `.html` files loaded via `HtmlService.createTemplateFromFile()`. Standardize all LLM calls to a strict JSON Schema output.
*   **Rejected/Deferred:**
    *   *Self-hosted GPU/Local Models:* Rejected due to requirement for free/low-cost serverless execution.
    *   *Direct Google Drive file mutation (Annotations):* Deferred to minimize privacy risk and scope creep.
    *   *FARA / Autonomous Agents:* Rejected. AI remains an extraction tool; human-in-the-loop is mandatory.

## 12. First Implementation Package Specification

**Target Branch:** `feature/phase1-core-stabilization`

**File-Level Change List:**
1.  `Code.gs`: Remove hardcoded `HtmlTemplates` strings. Update `openAssessmentStudio`, etc., to use `HtmlService.createTemplateFromFile()`. Include a spreadsheet backup step in `setupOrMigrateAssessmentSystem`.
2.  `AssessmentStudio.html`: Clean up and structure the file to be served directly, keeping `nearestGradeLetterClient` and `deriveGradeLocal` client-side as they currently represent the new logic.
3.  `CourseworkPicker.html`, `VoiceDictation.html`: Ensure they function as standalone loaded templates without duplicating code.
4.  `ClassBatchGrading.html`: Make sure it is extracted cleanly.
