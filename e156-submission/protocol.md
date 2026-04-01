Mahmood Ahmad
Tahir Heart Institute
author@example.com

Protocol: SubmissionCockpit - Rewrite Workbook and Editorial Release Ledger

This protocol describes a snapshot-first publication operations study for the C-drive portfolio. Eligible records are all indexed projects preserved in the bundled outputs from `ResearchConstellation`, `PortfolioOps`, `CitationWorkbench`, and `AuthorshipLedger`, with manual editorial state supplied by `data-source/rewrite-workbook.xlsx`. The primary estimand is the proportion of indexed projects carrying resolved submission-ready status at the time of build. Secondary outputs will report rewrite counts, E156 packaging state, GitHub readiness, GitHub Pages readiness, PDF galley readiness, and future Synthesis Journal upload readiness. The build process will mirror the workbook to CSV and JSON, refresh bundled upstream snapshots, emit `submission-cockpit.json`, and generate queue exports plus a static dashboard. Publication policy encoded in the tracker assumes E156 editorials, GitHub-hosted protocols, GitHub Pages dashboards, rerunnable code releases, and MIT licensing unless manually overridden. Anticipated limitations include workbook incompleteness, upstream snapshot lag, manual status errors, and the current absence of live journal API submission.

Outside Notes

Type: protocol
Primary estimand: proportion of indexed projects carrying resolved submission-ready status
App: SubmissionCockpit v0.1
Code: repository root, scripts/build_submission_cockpit.py, submission-cockpit.json, data-source/rewrite-workbook.xlsx
Date: 2026-04-01
Validation: DRAFT

References

1. Wilkinson MD, Dumontier M, Aalbersberg IJJ, et al. The FAIR Guiding Principles for scientific data management and stewardship. Sci Data. 2016;3:160018.
2. Sandve GK, Nekrutenko A, Taylor J, Hovig E. Ten simple rules for reproducible computational research. PLoS Comput Biol. 2013;9:e1003285.
3. Page MJ, McKenzie JE, Bossuyt PM, et al. The PRISMA 2020 statement. BMJ. 2021;372:n71.

AI Disclosure

This protocol was drafted from versioned local artifacts and deterministic build logic. AI was used as a drafting and implementation assistant under author supervision, with the author retaining responsibility for scope, methods, and reporting choices.
