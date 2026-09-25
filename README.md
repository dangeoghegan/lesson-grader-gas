# Lesson Grader

An Apps Script teacher workspace for rubric-based assessment across projects, with Google Classroom/Drive/Sheets as the data sources and teacher approval as the final grading step. The new web app supports **task-specific rubrics**, Classroom import or manual student entries, teacher grading, Gemini evidence proposals, mobile layouts, and versioned assessment history. The original bound Sheets Studio remains **Jewellery-only**.

**Start here:** [Deployment, Gemini key setup and safe testing](docs/IMPLEMENTATION.md) · [Original repository audit](docs/AUDIT.md)

**Local checks and mock UI preview:**

```bash
npm ci
npm run check && npm test
npm run preview          # port 4173, mock data only — not a Google deployment
```

No Gemini key is bundled. Configure it in the restricted deployed app's **Settings** (stored only in Script Properties) after reading the deployment checklist. The web app denies access until `LESSON_GRADER_ALLOWED_EMAILS` is configured. **Do not deploy over a live project without reconciling/exporting Dan's existing script and backing up the grading workbooks first.**
