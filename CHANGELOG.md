# Changelog

All notable changes to SystemUpdatePro will be documented in this file.

## [Unreleased]

- Added fail-closed ASUS, Acer, MSI, Surface, Framework, Panasonic, and Intel/AMD/NVIDIA GPU provider plans with signed local updater execution and public-source evidence.
- Added source-specific dependency readiness with bounded origin probes, verified SHA-256 content-addressed offline cache consumption, per-source timeout/proxy evidence, explicit WinGet machine/current-user/other-user scope results, protected package exclusion/pin/conflict planning, audited metered-network download policy, and deterministic local rollout cohort promotion evidence.
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

## Roadmap archive — 2026-08-10 — ROADMAP.md

<details>
<summary>Original roadmap snapshot</summary>

```markdown
# SystemUpdatePro Roadmap

Roadmap for SystemUpdatePro v4.1 - the self-healing enterprise update orchestrator covering OEM (Dell/Lenovo/HP), Windows Update, and Winget, with RMM-friendly exit codes and reporting.

## Planned Features

### OEM coverage
- ASUS MyASUS / Armoury Crate CLI driver pipeline
- Acer Care Center driver pipeline
- MSI Center / Dragon Center driver pipeline
- Generic Intel/AMD/NVIDIA GPU driver auto-update (bypass OEM when preferred, via public installer endpoints)
- Surface (Microsoft) firmware/driver pack detection + apply
- Framework laptops firmware update path
- Panasonic Toughbook driver pipeline

### Windows Update engine
- Feature-update auto-deferral with enterprise policy (stay on current LTS for N days)
- Driver-update allow/deny list with wildcard support
- Pre-staging support (download now, install at next reboot window)
- "Only critical + security" mode for maintenance-window short runs
- Microsoft Update Catalog fallback fetch when COM/PSWindowsUpdate both fail
- Windows ADMX policy auto-snapshot before touching WU components

### Winget & package managers
- Chocolatey and Scoop support as parallel upgrade sources (auto-detect installed)
- Per-package exclusion list (`winget-exclude.txt`) honored across runs
- Pin-version support (don't upgrade package X past version Y)
- Microsoft Store app upgrade leveraging `StoreEdgeFD` source
- Flatpak / Snap support if running WSL with GUI apps (stretch)

### Orchestration
- Maintenance window scheduler (respect Intune maintenance windows if detected)
- Staggered reboot coordination for clusters (don't reboot all machines at once)
- Power-management awareness (`powercfg` high-performance during run, restore after)
- Parallel OEM + WU path when safe (currently serial)
- Dry-run diff mode (what *would* be installed without any network calls beyond catalog check); first enforce a zero-persistent-mutation contract because current initialization, WinGet source refresh, and HP paths still write during `-DryRun`

### Reporting & integration
- Azure Monitor / Sentinel-friendly JSON schema for webhook payloads
- Prometheus textfile exporter output (`/ProgramData/.../metrics.prom`)
- Structured event log schema (XML event payload with parseable fields)
- Teams Adaptive Card webhook variant (not just plain Teams webhook)
- PSGallery publish as a module (`Install-Module SystemUpdatePro`)

### Safety & rollback
- Automatic restore point before each run (honor 24h throttle registry key)
- Rollback driver pipeline (DISM `/export-driver` + `/add-driver` on revert)
- Pre-run health check: CBS.log parse, DISM `/checkhealth`, sfc `/verifyonly`
- Post-run health check with same tools; fail run if health regressed

### CLI UX
- Color-aware console UI (ANSI-safe, fall back to plain text on legacy consoles)
- Progress bars per stage with `Write-Progress`
- Interactive mode (`-Interactive`) that asks before reboot

## Competitive Research

- **PatchMyPC Home Updater** - free consumer option, solid UX. Enterprise clones (PDQ Deploy, ManageEngine Patch Manager) add asset DB and fleet dashboards. SystemUpdatePro can differentiate as fleet-ready via RMM integration without their server.
- **WAU (Winget Auto Update)** - open-source Winget-only automation; SystemUpdatePro already dwarfs it, but mirror their exclusion-list format for drop-in compatibility.
- **Dell DCU / Lenovo LSUClient / HP IA** - already orchestrated. Add "update orchestration status" reporting back to the respective OEM dashboards where APIs allow.
- **PowerShell PSWindowsUpdate** - main WU module used; ensure compatibility with the latest version and fall back gracefully when module install fails in restricted environments.
- **Intune Update Rings** - not a replacement, but SystemUpdatePro should emit Intune-compatible compliance artifacts so it can coexist with an Intune-managed fleet.

## Nice-to-Haves

- MSI-wrapped installer for zero-touch deployment
- First-run wizard (interactive) that writes a persistent config for subsequent `-Unattended` runs
- Web dashboard (static HTML served from `\\share\SystemUpdatePro\dashboard`) with aggregated fleet reports
- PSScriptAnalyzer + Pester gate in CI
- Localization for German, French, Spanish, Japanese (enterprise customer asks)
- "Gold image validation" mode - compare a freshly-imaged machine against a stored baseline
- Optional SQLite-backed history instead of JSON for multi-machine consolidated reports

## Open-Source Research (Round 2)

### Related OSS Projects
- https://github.com/Romanitho/Winget-AutoUpdate — SYSTEM-context daily updates with allow/block lists, GPO, mods hook
- https://github.com/Sterbweise/winget-update — scheduled-task-oriented, smart detection, persistent exclusions
- https://github.com/microsoft/winget-cli — upstream Microsoft CLI + PowerShell module + COM API
- https://github.com/fire1ce/wingetup — JSON dump + Git-sync per hostname
- https://github.com/Kugane/winget — predefined-program installer with silent + MSStore support
- https://github.com/DellProSupport/DellCommandUpdate — Dell OEM baseline to cross-reference
- https://github.com/chocolatey/choco — alternative package manager; feature-parity targets
- https://github.com/mchoo1/HP-Image-Assistant-PowerShell — HPIA wrapper pattern
- https://github.com/Awcsh/patchmypcupdater — Patch My PC community wrappers

### Features to Borrow
- GPO/allowlist/blocklist loaded from URL/UNC with auto-refresh-on-newer (Winget-AutoUpdate)
- `_WAU-mods.ps1` hook convention — per-package pre/post script slot (Winget-AutoUpdate)
- SYSTEM-context scheduled task with logged-on-user notification via `msg.exe` / toast (Winget-AutoUpdate)
- JSON export of installed package state per hostname for fleet diffing (fire1ce/wingetup)
- `excluded_apps.txt` colocated with installer for airgapped overrides (Winget-AutoUpdate)
- PowerShell module companion (`Import-Module SystemUpdatePro`) publishing to PSGallery (microsoft/winget-cli pattern)
- `GH_TOKEN` / `GITHUB_TOKEN` auto-pickup for rate-limit bypass in CI (winget-cli PowerShell module)
- `--no-progress` / quiet mode switch for clean transcripts in RMM pipelines (winget-cli)
- Intune Win32 detection/install-script templates auto-generated per app (PowerShellIsFun pattern)
- Mod-hook for "app-with-config" installs (e.g., VSCode + settings.json) — borrow from WAU mods folder

### Patterns & Architectures Worth Studying
- Dual-mode execution: interactive GUI + silent SYSTEM task share same engine, diverge only at UI layer (Winget-AutoUpdate)
- External-list fetch with "update-only-if-newer" semantics — avoids clobbering air-gapped overrides
- Token-aware GitHub API client for release polling — essential for 1000+ endpoint MSP fleets
- MSI distribution with per-machine install + scheduled-task registration on install (WAU installer pattern)
- Per-app "mod script" override directory — lets field techs patch one-off machines without forking the tool

## Research-Driven Additions

### P0

### P1

- [ ] P1 — Add verified dependency cache and source-specific readiness
  Why: A generic Google probe neither predicts provider reachability nor supports proxy-restricted and air-gapped fleets; dependency download failure currently appears late inside privileged stages.
  Evidence: Verified — `SystemUpdatePro.ps1:531-553`, dependency installers; LSUClient offline repository/proxy documentation; WinGet source documentation.
  Touches: preflight, acquisition manifest, retry/proxy layer, WinGet/PSGallery/OEM adapters, CLI/config.
  Acceptance: Preflight reports readiness for each configured origin; an administrator can prefill a content-addressed cache containing the verified dependency manifest/artifacts; offline runs consume only matching cached artifacts; proxy and per-source timeout/failure reasons are preserved without contacting unrelated hosts.
  Complexity: L

- [ ] P1 — Model WinGet machine and user scopes explicitly
  Why: SYSTEM execution cannot reliably see or service all per-user packages, while invoking user-scoped installers from an elevated task can prompt, fail, or mutate the wrong profile.
  Evidence: Verified — Winget-AutoUpdate execution model; microsoft/winget-cli issues on SYSTEM context; Server Fault GPO/SYSTEM report; `SystemUpdatePro.ps1:1084-1147`.
  Touches: WinGet inventory/upgrade adapter, capability matrix, result schema, optional user-session boundary.
  Acceptance: Results distinguish machine, current-user, other-user, unavailable, and skipped packages; SYSTEM mode never claims success for unseen per-user scope; any user-session helper is non-elevating and separately authenticated; no hidden UAC prompt occurs in unattended mode.
  Complexity: L

### P2

- [ ] P2 — Enforce metered-network policy instead of warning only
  Why: The script detects network cost but proceeds with potentially large OEM, Windows, and application downloads.
  Evidence: Verified — `SystemUpdatePro.ps1:655-661`, `SystemUpdatePro.ps1:3230-3233`; Winget-AutoUpdate metered-network policy.
  Touches: preflight policy, provider download plans, CLI/config, result/report.
  Acceptance: Default policy blocks or defers large downloads on known metered links; an explicit audited override permits them; unknown network cost is reported distinctly; dry-run shows the decision without persistent mutation.
  Complexity: S

- [ ] P2 — Detect conflicting applications before package upgrades
  Why: Bulk `winget upgrade --all` can close, fail, or partially update applications that users are actively running.
  Evidence: Verified — `SystemUpdatePro.ps1:1084-1147`; Patch My PC conflicting-process notifications; PSAppDeployToolkit deferral/countdown model.
  Touches: WinGet plan/result adapter, package policy, optional non-elevating user notification boundary, unattended behavior.
  Acceptance: Policy maps package IDs to process/service conflicts and action (`skip`, `defer`, `close-with-deadline`); unattended default never force-closes an unknown process; conflicts and deferral deadlines are machine-readable; elevated engine remains isolated from user-session UI.
  Complexity: M

- [ ] P2 — Define progressive rollout policy and promotion evidence
  Why: MSP fleets need approved versions, pilot cohorts, bake time, and automatic halt thresholds before broad deployment; maintenance windows and reboot staggering alone do not prevent a bad update from propagating.
  Evidence: Verified — Action1 update rings, Patch My PC release controls, Windows Autopatch groups, Microsoft safe-deployment guidance, WAU issue 1179; existing Intune Update Rings and maintenance-window roadmap items.
  Touches: versioned policy/result schemas, package/OEM/WU selection, run IDs, webhook/export, configuration.
  Acceptance: An admin-authored policy assigns stable cohorts and approved package/KB/version rules with start/deadline, minimum success count/rate, and bake time; endpoints emit promotion evidence; a deterministic evaluator returns promote/hold/halt without requiring a hosted service; local emergency override is explicit and audited.
  Complexity: XL
```

</details>
