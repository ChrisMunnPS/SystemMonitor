# System Monitor

[![Version](https://img.shields.io/badge/version-1.0.0-6C63FF?style=flat-square)](https://github.com/ChrisMunnPS/SystemMonitor/releases)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?style=flat-square&logo=powershell)](https://microsoft.com/powershell)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-0078D4?style=flat-square&logo=windows)](https://github.com/ChrisMunnPS/SystemMonitor)
[![License](https://img.shields.io/badge/license-MIT-green?style=flat-square)](LICENSE)
[![WPF](https://img.shields.io/badge/UI-WPF%20%2F%20XAML-9B59B6?style=flat-square)](https://github.com/ChrisMunnPS/SystemMonitor)
[![Maintenance](https://img.shields.io/badge/maintained-yes-brightgreen?style=flat-square)](https://github.com/ChrisMunnPS/SystemMonitor/commits/main)

> A real-time WPF performance dashboard for Windows — built entirely in PowerShell.
> Monitor CPU, RAM, disk, network and processes with live charts, alerts, remote host support, and rich export options. No installation required.

---

## Quick Links

| | |
|---|---|
| 🌐 **Live Demo / Screenshots** | [ChrisMunnPS.github.io](https://ChrisMunnPS.github.io) |
| 📄 **Latest HTML Report Sample** | [View sample report](#) *(placeholder)* |
| 📊 **Sample CSV Export** | [Download sample CSV](#) *(placeholder)* |
| 🐛 **Report a Bug** | [Open an issue](https://github.com/ChrisMunnPS/SystemMonitor/issues) |
| 💡 **Request a Feature** | [Start a discussion](https://github.com/ChrisMunnPS/SystemMonitor/discussions) |

---

## Executive Summary

**System Monitor** is a single-file PowerShell script that launches a dark-themed WPF dashboard giving you instant, live insight into your Windows machine — or any remote Windows host on your network.

It replaces the need to juggle Task Manager, Resource Monitor, and Event Viewer for day-to-day performance checks. Everything is visible in one window: live KPI cards with colour-coded thresholds, donut charts per drive, a filterable process list with end-task capability, a one-click tool strip for common Windows utilities, and rich export options for reporting.

No installation, no dependencies beyond PowerShell 5.1 and .NET 4.x (both built into Windows 10/11). Just run the script.

---

## Table of Contents

- [Features](#features)
- [Screenshots](#screenshots)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Parameters](#parameters)
- [Remote Host Monitoring](#remote-host-monitoring)
- [Exports](#exports)
- [Tool Strip](#tool-strip)
- [Thresholds & Alerts](#thresholds--alerts)
- [Performance Notes](#performance-notes)
- [Versioning](#versioning)
- [Author](#author)
- [License](#license)

---

## Features

| Category | Details |
|---|---|
| **Live KPI Cards** | CPU %, RAM %, Primary Disk %, Network IN/OUT — updated every N seconds |
| **Drive Donut Charts** | One animated donut per drive showing used %, free space, and total capacity |
| **Process Manager** | Sortable by CPU, Memory or Name; live CPU % delta; filter by name or PID; right-click End Task / Open File Location |
| **Remote Monitoring** | Connect to any Windows host via WinRM/CIM — all metrics, process list and End Task work over the network |
| **Battery Status** | Charge %, status (charging/discharging) and estimated run time — auto-hidden on desktops |
| **Network Detail** | Primary adapter name, IP address, link speed, and async gateway ping latency |
| **Alert System** | Configurable thresholds for CPU, RAM and Disk — visual indicator, alert log, and persistent log file in `%TEMP%` |
| **10-Second Rolling Stats** | Status bar shows rolling average and peak for all metrics over the last 10 seconds |
| **Tool Strip** | One-click launch of Disk Cleanup, Reliability Monitor, Wi-Fi Report, Services, Resource Monitor, Event Viewer, Device Manager, Disk Management, System Info, Performance Monitor, Task Manager, System Config |
| **Export: CSV** | Full session history (up to 1000 samples) exported to CSV |
| **Export: Markdown** | Summary stats, alert log and sample table in clean Markdown |
| **Export: HTML** | Fully charted dark-theme report with 7 Chart.js visualisations |
| **Auto-Export** | Schedule HTML export every N minutes to `%TEMP%` automatically |
| **Auto-Refresh Toggle** | Checkbox to freeze the UI (useful when selecting / ending processes) + manual Refresh Now button |

---

## Screenshots

> *(Add screenshots here — drag images into this section on GitHub)*

---

## Requirements

| Requirement | Minimum | Notes |
|---|---|---|
| Windows | 10 (build 18362+) or 11 | WPF requires Windows; tool strip buttons require standard Windows tools |
| PowerShell | 5.1 | Built into Windows 10/11. PowerShell 7 is supported (auto-relaunches in 5.1 STA mode) |
| .NET Framework | 4.7.2+ | Ships with Windows 10/11 |
| WinRM (remote only) | Enabled on target | Run `Enable-PSRemoting -Force` as Administrator on the remote machine |

---

## Installation

**Option 1 — Clone**
```powershell
git clone https://github.com/ChrisMunnPS/SystemMonitor.git
cd SystemMonitor
```

**Option 2 — Download single file**

Download [`SystemMonitor.ps1`](https://github.com/ChrisMunnPS/SystemMonitor/raw/main/SystemMonitor.ps1) directly — no other files needed.

**Unblock if downloaded from the web:**
```powershell
Unblock-File -Path .\SystemMonitor.ps1
```

---

## Usage

```powershell
# Default — 2 second refresh, standard thresholds
.\SystemMonitor.ps1

# Custom refresh rate and thresholds
.\SystemMonitor.ps1 -RefreshSeconds 3 -CpuThreshold 80 -RamThreshold 85 -DiskThreshold 95

# From an MTA host (VS Code, pwsh 7) — the script auto-relaunches in STA mode
# No special syntax needed; just run it normally
```

If your execution policy blocks the script:
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\SystemMonitor.ps1
```

---

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-RefreshSeconds` | int (1–60) | `2` | Auto-refresh interval in seconds |
| `-CpuThreshold` | int (1–100) | `75` | CPU % at which alerts trigger |
| `-RamThreshold` | int (1–100) | `80` | RAM % at which alerts trigger |
| `-DiskThreshold` | int (1–100) | `90` | Disk % at which alerts trigger |

Thresholds can also be changed at runtime via the **Thresholds** button without restarting.

---

## Remote Host Monitoring

1. On the **target machine**, open an elevated PowerShell and run:
   ```powershell
   Enable-PSRemoting -Force
   ```
2. In System Monitor, type the hostname or IP address into the **Remote Host** box and press **Enter** or click **Connect**.
3. The status indicator turns green and all metrics switch to the remote machine.
4. **End Task** works remotely via CIM — requires the connecting account to have admin rights on the target.
5. Click **Disconnect** to return to local monitoring.

> **Tip:** Hostnames are remembered in the dropdown for the session.

---

## Exports

### CSV
Full session history exported as a standard CSV — compatible with Excel, Power BI, or any data tool.

### Markdown
Produces a clean `.md` file with:
- Session summary stats table (average and peak per metric)
- Total network transferred
- Alert log
- Full sample data table

### HTML Report
A self-contained dark-theme HTML file with **7 Chart.js visualisations**:

| # | Chart Type | Shows |
|---|---|---|
| 1 | Multi-series line | CPU / RAM / Disk % over time |
| 2 | Stacked area | Network IN + OUT over time |
| 3 | Grouped bar (last 40) | Resource balance snapshot |
| 4 | Side-by-side columns (last 40) | Net IN vs Net OUT |
| 5 | Scatter | CPU % vs Net IN (correlation) |
| 6 | Scatter | CPU % vs RAM % (correlation) |
| 7 | Heatmap bar | Resource intensity over time |

[View sample HTML report](#) *(placeholder — replace with a link to a hosted sample)*

### Auto-Export
Set an interval in **Thresholds → Auto HTML (mins)** to automatically write an HTML report to `%TEMP%` on a schedule.

---

## Tool Strip

One-click launch strip across the top of the dashboard:

| Button | Launches |
|---|---|
| Disk Cleanup | `cleanmgr` |
| Reliability | Reliability History (`perfmon /rel`) |
| WiFi Report | Generates and opens `wlan-report-latest.html` |
| Services | `services.msc` |
| Resource Mon | `resmon` |
| Event Viewer | `eventvwr` |
| Device Manager | `devmgmt.msc` |
| Disk Management | `diskmgmt.msc` |
| System Info | `msinfo32` |
| Perf Monitor | `perfmon` |
| Task Manager | `taskmgr` |
| System Config | `msconfig` |

---

## Thresholds & Alerts

- Set per-metric thresholds via the **Thresholds** button or at launch via parameters.
- When a threshold is breached:
  - A `[!]` indicator appears on the KPI card
  - An entry is added to the **Alert Log** panel
  - The alert is appended to `%TEMP%\SystemMonitor_alerts.log` (persistent across sessions)
  - The status bar shows a 10-second rolling average with peak

---

## Performance Notes

The dashboard is optimised to keep the UI thread responsive even under load:

- **CPU %** — uses `System.Diagnostics.PerformanceCounter` locally; `Win32_Processor.LoadPercentage` remotely
- **Per-process CPU %** — delta calculated from `KernelModeTime + UserModeTime` (same method as Task Manager)
- **Process `.Responding`** — bulk `Get-Process` once per tick into a hashtable (not per-process)
- **Ping** — fully async (`SendPingAsync`) — never blocks the UI thread
- **Adapter / battery info** — cached and refreshed every 30–60 seconds
- **Drive charts** — only rebuilt when drive usage values actually change

---

## Versioning

This project follows [Semantic Versioning 2.0.0](https://semver.org/):

```
MAJOR.MINOR.PATCH
```

| Increment | When |
|---|---|
| **MAJOR** | Breaking changes, major UI overhaul, removal of features |
| **MINOR** | New features added in a backwards-compatible manner |
| **PATCH** | Bug fixes, performance improvements, minor UI tweaks |

See [CHANGELOG](CHANGELOG.md) for release history.

### Current Version: `v1.0.0`

Initial public release featuring:
- Live KPI dashboard (CPU, RAM, Disk, Network)
- Drive donut charts
- Process manager with End Task
- Remote host monitoring via WinRM/CIM
- Battery, ping, adapter detail
- Alert system with persistent log
- Tool launch strip (12 tools)
- Export to CSV, Markdown, and charted HTML
- Auto-export scheduling

---

## Author

**Christopher Munn**

[![GitHub](https://img.shields.io/badge/GitHub-ChrisMunnPS-181717?style=flat-square&logo=github)](https://github.com/ChrisMunnPS/SystemMonitor)
[![Website](https://img.shields.io/badge/Website-ChrisMunnPS.github.io-10B981?style=flat-square&logo=googlechrome)](https://ChrisMunnPS.github.io)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-chrismunn-0A66C2?style=flat-square&logo=linkedin)](https://www.linkedin.com/in/chrismunn)

---

## License

MIT License — see [LICENSE](LICENSE) for details.

```
MIT License

Copyright (c) 2026 Christopher Munn

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
```
