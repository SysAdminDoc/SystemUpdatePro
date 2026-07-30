# Changelog

All notable changes to SystemUpdatePro will be documented in this file.

## [Unreleased]

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
