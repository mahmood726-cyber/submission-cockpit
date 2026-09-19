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

## Limits (read before quoting a number)

- **The committed `submission-cockpit.json` (generated 2026-04-12) reports `workbookRewriteCount: 457, workbookValidRewriteCount: 0`. Both numbers are wrong.** The workbook parser captured the trailing `SUBMITTED: [ ]` state line as part of every rewrite, so each rewrite had 8 "sentences" and a blank line, and could never equal its CURRENT BODY. Re-parsed with the state line excluded (`tests/test_workbook_parser.py`), the bundled snapshot holds 469 entries: **236 real rewrites, 102 rewrites identical to the current body, 131 blank; 199 of 236 (84%) pass the 7-sentence validator; 0 marked SUBMITTED.** A rebuild from the bundled snapshots (`SUBMISSION_COCKPIT_FROM_SNAPSHOTS=1`) gives `rewrittenCount 190, workbookValidRewriteCount 190, readyForGitHubCount 1` for the 466 tracked rows; the committed artefacts have deliberately not been regenerated on this branch because the rebuild rewrites the manual tracker mirror (`data-source/rewrite-tracker.*`, `rewrite-workbook.xlsx`) wholesale.
- **The builder cannot run from a public clone by default.** `sync_snapshots()` copies from sibling directories on the author's machine (`../ResearchConstellation`, `../PortfolioOps`, `C:\E156`) and raises if they are absent. Set `SUBMISSION_COCKPIT_FROM_SNAPSHOTS=1` to build from `data-source/*.snapshot.*` instead. Other reads of local paths (project metadata, `repo_mapping.json`) still fall back silently when absent, so an offline build is a snapshot build, not a reproduction of the author's machine.
- **The validator's 100-word floor is this tool's choice.** The E156 contract is "exactly 7 sentences, at most 156 words"; the floor is stricter than the specification.
- **`SUBMITTED: [x]` was never read** before this branch; the `submitted` field is now parsed but not yet wired into readiness scoring or the tracker mirror.
- Windows paths in the data are the registry's identity field and are intentionally not portable.
