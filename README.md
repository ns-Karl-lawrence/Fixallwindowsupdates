# FixAllWindowsUpdates

[![PowerShell CI](https://github.com/KarlLawrence/FixAllWindowsUpdates/actions/workflows/powershell-ci.yml/badge.svg)](https://github.com/KarlLawrence/FixAllWindowsUpdates/actions/workflows/powershell-ci.yml)

A single PowerShell script (`Fixallwindowsupdates.ps1`) that performs end-to-end Windows Update recovery: health checks, optional repair routines, comprehensive logging, and interactive or unattended installation of available updates with retry and verification logic.@Fixallwindowsupdates.ps1#2-114@Fixallwindowsupdates.ps1#121-134

> **Heads up:** Update the badge slug above if you publish the script under a different GitHub organization or repo name.

## Key capabilities

- Pre-flight validations for disk space, networking, pending reboots, Windows Update Agent (WUA) health, and optional extended connectivity tests.@Fixallwindowsupdates.ps1#2-114@Fixallwindowsupdates.ps1#1507-1684
- Automatic system restore point creation before changes (opt-out via `-SkipSystemRestore`).@Fixallwindowsupdates.ps1#2-114@Fixallwindowsupdates.ps1#3571-3578
- Tiered repair routines that reset services, caches, security descriptors, DLL registrations, BITS queues, DISM/SFC (when `-AggressiveRepair`), and Windows Update policy artifacts.@Fixallwindowsupdates.ps1#65-104@Fixallwindowsupdates.ps1#2770-2836
- Update discovery via the Windows Update COM APIs with category breakdowns, exclusive-update awareness, and interactive selection when auto-install is disabled.@Fixallwindowsupdates.ps1#2838-2992@Fixallwindowsupdates.ps1#3600-3650
- Batch installs with retry logic, optional per-update retries, fallback repairs, and re-verification scans to confirm success.@Fixallwindowsupdates.ps1#2994-3492@Fixallwindowsupdates.ps1#3600-3678
- Rich logging with correlation IDs, phase timings, log rotation, Windows Event Log writes for errors, transcripts, and optional CSV recommendations.@Fixallwindowsupdates.ps1#147-213@Fixallwindowsupdates.ps1#3376-3416

## Requirements

- Windows 10/11 or Windows Server 2016+ with Windows Update enabled.@Fixallwindowsupdates.ps1#106-117
- PowerShell 5.1+ running **as Administrator** (UAC elevation required).@Fixallwindowsupdates.ps1#106-114@Fixallwindowsupdates.ps1#3509-3535
- Internet access for Microsoft Update endpoints, or a reachable WSUS server if one is configured.@Fixallwindowsupdates.ps1#1579-1684
- Enough free disk space (default threshold: 10 GB) for staging updates and repair caches.@Fixallwindowsupdates.ps1#1553-1576

## Getting started

1. Download or this repository.
2. Open an elevated PowerShell console (Run as Administrator).
3. Unblock the script if it was downloaded from the internet:

   ```powershell
   Unblock-File .\Fixallwindowsupdates.ps1
   ```

4. Run one of the sample commands below.

### Common recipes

- **Standard repair + auto install (default behavior):**

  ```powershell
  .\Fixallwindowsupdates.ps1 -AutoInstall
  ```

- **Skip repairs, only scan/install updates:**

  ```powershell
  .\Fixallwindowsupdates.ps1 -SkipFix -AutoInstall
  ```

- **Aggressive repair with extended retries:**

  ```powershell
  .\Fixallwindowsupdates.ps1 -AggressiveRepair -MaxRetries 5
  ```

- **Dry-run style scan with manual approval:**

  ```powershell
  .\Fixallwindowsupdates.ps1 -SkipSystemRestore -SkipPreflightChecks
  ```

Use `Get-Help .\Fixallwindowsupdates.ps1 -Detailed` for inline documentation and the warnings table embedded at the top of the script.@Fixallwindowsupdates.ps1#2-114

## Parameter reference

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `SkipFix` | switch | `False` | Bypass the repair phase and proceed directly to scanning/installing updates.@Fixallwindowsupdates.ps1#16-18@Fixallwindowsupdates.ps1#3579-3599 |
| `SkipPreflightChecks` | switch | `False` | Skips disk, network, pending reboot, and WUA health checks.@Fixallwindowsupdates.ps1#19-21@Fixallwindowsupdates.ps1#3548-3569 |
| `SkipSystemRestore` | switch | `False` | Prevents automatic system restore point creation.@Fixallwindowsupdates.ps1#22-24@Fixallwindowsupdates.ps1#3571-3578 |
| `AutoInstall` | switch | `True` | Automatically install all detected updates without prompting.@Fixallwindowsupdates.ps1#25-27@Fixallwindowsupdates.ps1#3611-3618 |
| `ForceAutoInstall` | switch | `False` | Overrides safety prompts to continue even when warnings exist.@Fixallwindowsupdates.ps1#28-30@Fixallwindowsupdates.ps1#3543-3564 |
| `AggressiveRepair` | switch | `False` | Adds DISM, SFC, and Windows Update policy resets to the repair phase.@Fixallwindowsupdates.ps1#31-33@Fixallwindowsupdates.ps1#2770-2836 |
| `TestConnectivity` | switch | `False` | Runs extended Microsoft/WSUS endpoint tests during preflight.@Fixallwindowsupdates.ps1#34-36@Fixallwindowsupdates.ps1#1579-1684 |
| `MaxRetries` | int | `3` | Maximum attempts for failed downloads/installs.@Fixallwindowsupdates.ps1#37-38@Fixallwindowsupdates.ps1#2994-3304 |
| `LogPath` | string | `C:\Logs\WindowsUpdate_<timestamp>.log` | File path for structured log output (directories auto-created).@Fixallwindowsupdates.ps1#40-41@Fixallwindowsupdates.ps1#147-213 |
| `TranscriptPath` | string | `<LogPath>_transcript.log` | Optional PowerShell transcript path (autogenerated if omitted).@Fixallwindowsupdates.ps1#43-45@Fixallwindowsupdates.ps1#3498-3508 |
| `RecommendationsCsvPath` | string | `C:\Logs\WindowsUpdate_Recommendations_<timestamp>.csv` | CSV-export path for update recommendations and outcomes.@Fixallwindowsupdates.ps1#46-48@Fixallwindowsupdates.ps1#3409-3416 |

## Execution flow

1. **Pre-flight checks** – Validates prerequisites, optionally prompting when issues are found (unless `-SkipPreflightChecks`).@Fixallwindowsupdates.ps1#3548-3569
2. **System restore point** – Creates a restore point for safe rollback (unless `-SkipSystemRestore`).@Fixallwindowsupdates.ps1#3571-3578
3. **Repair phase** – Optional multi-step repair routine, elevated to "aggressive" when requested.@Fixallwindowsupdates.ps1#3579-3599@Fixallwindowsupdates.ps1#2770-2836
4. **Scan** – Queries the Windows Update catalog, categorizes updates, flags exclusive installs.@Fixallwindowsupdates.ps1#2838-2954@Fixallwindowsupdates.ps1#3600-3609
5. **Selection** – Auto-selects all updates or prompts for manual choices based on install mode.@Fixallwindowsupdates.ps1#3611-3650
6. **Installation & retries** – Downloads/installs in batches, falls back to per-update retries and repair loops if needed.@Fixallwindowsupdates.ps1#2994-3492
7. **Verification & summary** – Re-scan, surface remaining updates, and print a structured summary with reboot guidance.@Fixallwindowsupdates.ps1#3307-3416

## Logging & artifacts

- **Primary log** (`LogPath`): Timestamped entries including severity, correlation IDs, phase timings, and rotation when >10 MB.@Fixallwindowsupdates.ps1#147-213
- **PowerShell transcript** (`TranscriptPath`): Full console history for auditing/troubleshooting.@Fixallwindowsupdates.ps1#3498-3508
- **Recommendations CSV** (`RecommendationsCsvPath`): Optional report of update actions and tips.@Fixallwindowsupdates.ps1#46-48@Fixallwindowsupdates.ps1#3409-3416
- **Windows Event Log**: Errors mirrored to *Application* log under source `WindowsUpdateScript`.@Fixallwindowsupdates.ps1#195-207

## Safety considerations

The script manipulates critical Windows Update components (services, folders, DLL registrations, DISM/SFC, registry keys, and ACLs). Review the RISKS table near the top of the script before running in production or on sensitive endpoints.@Fixallwindowsupdates.ps1#65-104

## Troubleshooting tips

- Review `C:\Logs\WindowsUpdate_*.log` and the PowerShell transcript for context.
- Inspect `CBS.log` / `WindowsUpdate.log` (captured in the verification phase) for low-level servicing stack issues.@Fixallwindowsupdates.ps1#3330-3366
- Use `-AggressiveRepair` for stubborn corruption, then re-run with `-AutoInstall`.
- When WSUS is in play, ensure the server URL resolves and that required ports are open; `-TestConnectivity` can help validate endpoints.@Fixallwindowsupdates.ps1#1579-1684

## Automation & CI

A GitHub Actions workflow (`.github/workflows/powershell-ci.yml`) is included to lint the script with PSScriptAnalyzer on every push/PR. It also contains a conditional Pester step that auto-runs whenever `Tests/**/*.Tests.ps1` files are added. To run the same checks locally:

```powershell
Install-Module PSScriptAnalyzer -Scope CurrentUser -Force
Invoke-ScriptAnalyzer -Path .\Fixallwindowsupdates.ps1 -Settings PSGallery\PSScriptAnalyzerSettings.psd1
```

You can expand the workflow with integration tests or signed-execution checks if you add Pester tests or a test harness later.

## Contributing

Feedback and contributions are welcome! Please open an issue with:

1. The command line you used (parameters and switches)
2. Relevant log excerpts (redact sensitive info)
3. Windows build number and whether WSUS/proxy policies are applied

## License

This project is licensed under the [MIT License](LICENSE).
