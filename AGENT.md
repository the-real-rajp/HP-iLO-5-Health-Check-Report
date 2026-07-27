# HP iLO 5 Health Check Report: Maintainer Guide

## Scope and runtime

- Support PowerShell 7.0 or later. Windows PowerShell 5.1 is unsupported and
  must not be documented as a supported runtime.
- Keep the implementation self-contained; no external PowerShell modules are
  required.
- The script uses Redfish session authentication and must never write
  passwords, session tokens, or authorization headers to the console, logs,
  reports, or tests.

## Report behavior

- `New-OpenXmlHealthReport` is the default report writer. The legacy Word COM
  writer remains opt-in through `HP_ILO_USE_WORD_COM=1`.
- Preserve the existing Winslow Tech Group header, footer, narrow margins,
  assessment wording, and table filtering unless a requested change says
  otherwise.
- Use the shared assessment helpers for severity decisions. A Security Status
  of `Risk` is `CRITICAL`; an ignored overall or individual security finding is
  `WARNING`.
- Keep unavailable, `Unknown`, `Absent`, and `N/A` evidence out of reports in
  accordance with the existing report-record filters.

## Progress and logging

- Use `Write-ReportProgress` for user-facing collection status. Parent groups
  have no indent; child steps use `-Indent 1` and display as `  - Step`.
- Use `Write-ReportLog` for diagnostic events. The default timestamped log is
  stored in `logs/` and must remain free of credentials and tokens.
- Collection failures should be recorded as collection notes when the report
  can still be generated. Do not silently hide failed data collection.

## Validation and sample updates

- Run `pwsh -NoProfile -File ./tests/Smoke.Tests.ps1` after script changes.
- When changing report layout or report data, regenerate the checked-in sample
  with `pwsh -NoProfile -File ./examples/Generate-SampleReport.ps1`.
- Render and inspect `examples/sample-health-report.docx` after regenerating
  it when LibreOffice is available. If rendering is unavailable, perform
  structural DOCX checks and record that limitation in the handoff.
- Keep smoke tests offline. They must not contact an iLO or start Microsoft
  Word.

## Repository hygiene

- Do not commit generated run logs, customer reports, credentials, or real
  infrastructure data.
- Update `README.md`, the sample report, and smoke tests whenever a
  user-visible feature or documented behavior changes.
