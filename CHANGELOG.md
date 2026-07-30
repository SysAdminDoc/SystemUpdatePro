# Changelog

All notable changes to SystemUpdatePro will be documented in this file.

## [Unreleased]

- Added webhook payload schema v2 with deterministic idempotency keys, run/timestamp/stage correlation, report evidence URIs, no-redirect HTTPS transport, bounded transient/429 retry with `Retry-After`, durable per-attempt/terminal records, Slack and legacy connector formatting, and Teams Workflow Adaptive Cards generated from the same contract.
- Replaced raw webhook command-line input with validated environment/protected-file secret references, added pre-initialization range/path/HTTPS validation, advanced continuation state to schema v5 with v3/v4 migration, retained read-only history as an unelevated early command, and extended redaction across console, transcript, exceptions, reports, history, and diagnostic bundles.
- Added a standalone, bounded diagnostic/recovery bundle with a schema-versioned SHA-256 manifest, latest run and policy data, runtime/provider capability inventory, redacted transcript and OEM output, Windows Update/USO/CBS/DISM evidence, mutation-journal status, explicit collector errors, atomic protected publication, and deterministic PowerShell 5.1/7 ZIP paths.
- Added a protected local evidence store with write-through atomic replacement, last-known-good recovery, corrupt-file quarantine, forward state/history migration, verified SYSTEM/Administrators ACLs, recursive secret/serial redaction, and exact age/size retention across logs, reports, OEM output, journals, quarantine files, and driver backups.
- Added a machine-readable platform/provider capability matrix with tested Windows client/Server/Core, x86/x64/ARM64, PowerShell 5.1/7, administrator/SYSTEM, manufacturer, and installed-version gates; unsupported stages now skip explicitly and capability evidence persists across continuation, reports, history, and webhooks.
- Added protected, atomic, run-scoped mutation journals with startup crash recovery and verified reverse-order restoration for WSUS policy, service state, update-cache swaps, cleanmgr flags, Dell services, and continuation tasks; Windows Update repair now diagnoses first and no longer resets Winsock/WinHTTP, re-registers DLLs, or deletes cache contents in place.
- Added a single elevated-dependency acquisition manifest with approved HTTPS origins, redirect constraints, architecture/version floors, exact package hashes, Authenticode publisher checks, verified module staging, WinGet/Dell source contracts, HPIA 5.3.6, Dell Inventory Collector 13.8.0+ enforcement, and persisted report/history/webhook provenance without weakening PowerShell Gallery trust.
- Made firmware installation fail closed across Dell, Lenovo, and HP with tri-state disk/power/BitLocker checks, a configurable 50% charge floor, provider/model applicability scans, one shared firmware filter, actionable block reasons, and no `-Force` bypass for unknown safety state.
- Made component cleanup reversible by default, moved irreversible DISM `/ResetBase` behind `-ResetComponentBase`, exposed rollback impact in dry runs and reports, and now restores temporary Disk Cleanup registry flags while reporting partial failures.
- Rebuilt post-reboot continuation as a PowerShell 5.1-compatible state machine with atomic protected state, full parameter/run restoration, resumable stage cursors, bounded attempts, corrupt-state quarantine, registration rollback, and terminal task cleanup.
- Normalized OEM, Windows Update, and WinGet outcomes into a schema-versioned stage/item result contract with truthful totals, provider exit/HRESULT evidence, aligned process exit codes, and exactly-once terminal report/history/Event Log/webhook attempts.
- Redesigned HTML reports as responsive operations dashboards with encoded data, clear update-channel states, device inventory, exception guidance, and print styles.

## [v4.1.0] - %Y->- (HEAD -> main, tag: v4.1.0, origin/main)

- Added: Add project icon to README
- v4.1.0 - DryRun mode, HTML reports, webhook notifications, driver backup, update history
- Initial commit - SystemUpdatePro
