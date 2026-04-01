# SubmissionCockpit

SubmissionCockpit is the execution layer for the C-drive research portfolio.

It turns the existing portfolio, citation, authorship, readiness outputs, and the live `C:\E156\rewrite-workbook.txt` tracker into one publication ledger centered on `E156` editorials, GitHub-hosted protocols, GitHub Pages dashboards, MIT-licensed code, and future Synthesis Journal uploads.

## Why this exists

The current portfolio already has:

- a project atlas in `ResearchConstellation`
- operational signals in `PortfolioOps`
- citation packets in `CitationWorkbench`
- deposit and contributor scaffolds in `AuthorshipLedger`
- an `E156` framework in `E156-framework`

What it still lacked was one canonical tracker that answers:

- which papers have been rewritten
- which projects should become new `E156` editorials next
- which repos still need MIT licensing, GitHub sync, or GitHub Pages
- which projects still need a protocol, dashboard, PDF galley, or journal upload

## What it does

- copies bundled snapshots from the upstream portfolio projects
- creates and maintains a canonical rewrite workbook in `data-source/rewrite-workbook.xlsx`
- mirrors the workbook to `rewrite-tracker.csv` and `rewrite-tracker.json` for reproducibility
- merges manual rewrite state with portfolio, citation, authorship, and deposit signals
- generates per-project publication manifests and queue exports
- serves a static GitHub Pages dashboard from `index.html`
- ships its own `E156` paper and protocol bundle

## Tracker workflow

The canonical rewrite source is:

`C:\E156\rewrite-workbook.txt`

That text workbook already drives the external `E156` rewrite lane and is parsed directly by SubmissionCockpit.

SubmissionCockpit also maintains a local mirror workbook for release-state fields such as GitHub, Pages, PDF, and journal upload state:

`data-source/rewrite-workbook.xlsx`

Edit that workbook, then rebuild:

```bash
python3 /mnt/c/Users/user/SubmissionCockpit/scripts/build_submission_cockpit.py
```

The build will:

- refresh upstream snapshot copies
- snapshot `C:\E156\rewrite-workbook.txt` plus its maintenance reports
- parse live rewrite entries from the external workbook
- read workbook values
- preserve your manual status fields
- append any newly discovered projects
- rewrite the CSV and JSON mirrors
- regenerate queue outputs and the dashboard payload

## Outputs

- `submission-cockpit.json` - full merged publication ledger
- `data.json` and `data.js` - static dashboard payloads
- `exports/rewrite-queue.json` - projects still needing rewrite completion
- `exports/build-queue.json` - rewritten projects still missing paper/protocol/dashboard packaging
- `exports/github-queue.json` - projects ready to push to GitHub
- `exports/pages-queue.json` - projects ready for GitHub Pages publication
- `exports/synthesis-queue.json` - projects ready for PDF or Synthesis upload steps
- `exports/project-manifests/` - one manifest per tracked project

## Publication policy encoded here

- paper type: `E156` editorial
- protocol location: GitHub
- dashboard location: GitHub Pages
- code expectation: rerunnable and easy to copy
- license target: MIT unless you explicitly override it
- future journal path: Synthesis Journal editorial upload with PDF galley support

## Future hook

This scaffold reserves a future Synthesis Journal integration path. When you provide credentials later, the natural environment variable to wire in is:

`SYNTHESIS_JOURNAL_API_KEY`

The current build tracks readiness for that upload step but does not attempt API submission yet.

## Existing E156 automation

The external `C:\E156` repo already contains:

- `scripts/apply_rewrites.py` for validating and applying workbook rewrites
- `scripts/daily_sync.bat` for the current daily sync lane
- `audit-report.json`, `verification-report.json`, and `maintenance-report.json` for health checks

SubmissionCockpit snapshots those artifacts so the publication dashboard can stay aligned with the existing rewrite pipeline rather than replacing it.
