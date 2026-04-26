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
- Dry-run diff mode (what *would* be installed without any network calls beyond catalog check)

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
