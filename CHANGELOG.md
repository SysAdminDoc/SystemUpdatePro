# Changelog

All notable changes to SystemUpdatePro will be documented in this file.

## [Unreleased]

- Made firmware installation fail closed across Dell, Lenovo, and HP with tri-state disk/power/BitLocker checks, a configurable 50% charge floor, provider/model applicability scans, one shared firmware filter, actionable block reasons, and no `-Force` bypass for unknown safety state.
- Made component cleanup reversible by default, moved irreversible DISM `/ResetBase` behind `-ResetComponentBase`, exposed rollback impact in dry runs and reports, and now restores temporary Disk Cleanup registry flags while reporting partial failures.
- Rebuilt post-reboot continuation as a PowerShell 5.1-compatible state machine with atomic protected state, full parameter/run restoration, resumable stage cursors, bounded attempts, corrupt-state quarantine, registration rollback, and terminal task cleanup.
- Normalized OEM, Windows Update, and WinGet outcomes into a schema-versioned stage/item result contract with truthful totals, provider exit/HRESULT evidence, aligned process exit codes, and exactly-once terminal report/history/Event Log/webhook attempts.
- Redesigned HTML reports as responsive operations dashboards with encoded data, clear update-channel states, device inventory, exception guidance, and print styles.

## [v4.1.0] - %Y->- (HEAD -> main, tag: v4.1.0, origin/main)

- Added: Add project icon to README
- v4.1.0 - DryRun mode, HTML reports, webhook notifications, driver backup, update history
- Initial commit - SystemUpdatePro
