# Lesson Grader

An Apps Script teacher workspace for rubric-based assessment across projects, with Google Classroom/Drive/Sheets as the data sources and teacher approval as the final grading step. The new web app supports **task-specific rubrics**, Classroom import or manual student entries, teacher grading, Gemini evidence proposals, mobile layouts, and versioned assessment history. The original bound Sheets Studio remains **Jewellery-only**.

**Start here:** [Deployment, Gemini key setup and safe testing](docs/IMPLEMENTATION.md) · [Original repository audit](docs/AUDIT.md) · [GitHub Pages launcher](docs/LAUNCHER.md)

**Open it from GitHub:** `site/index.html` is a static launcher page for GitHub Pages. It holds no deployment
address — the teacher pastes the Apps Script **web app URL** into the page's own **Settings**, it is saved in the
browser, and the deployment (with its own sign-in and allowlist) loads inside the page. See
[docs/LAUNCHER.md](docs/LAUNCHER.md); `.github/workflows/pages.yml` publishes it.

**Local checks and mock UI preview:**

```bash
npm ci
npm run check && npm test
npm run preview          # port 4173, mock data only — not a Google deployment
                         #   /          mock web app UI
                         #   /launcher  the static GitHub Pages launcher
```

No Gemini key is bundled. Configure it in the restricted deployed app's **Settings** (stored only in Script Properties) after reading the deployment checklist. The web app denies access until `LESSON_GRADER_ALLOWED_EMAILS` is configured. **Do not deploy over a live project without reconciling/exporting Dan's existing script and backing up the grading workbooks first.**
