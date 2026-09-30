# Overview

A single PowerShell script that reads the Citrix Monitor Service OData API and produces a self-contained HTML report of unique users and concurrent sessions.

It is **read-only** — every call is a GET, and it changes nothing in your environment. It installs nothing: one `.ps1`, no modules, no agents, Windows PowerShell 5.1 baseline. It states what was measured and draws no conclusions.

**Documentation is three files.** This one: what it is and how to run it.
- `RUNBOOK`: the operational guide, including combining several sites, troubleshooting and the FAQ.
- `REFERENCE`: permissions, exactly what data is read and written, the design decisions and what was measured to justify them.

---

# What you need

* A **Delivery Controller** hostname — the Monitor Service runs there.
* Whether that controller publishes Monitor over **HTTPS or HTTP**. The default in this tool is **HTTPS** but the default when Citrix is installed is **HTTP**.
* An account with **read access to Citrix Monitor data**. Full administrator is not required.
* For Citrix Cloud instead: customer ID, service principal ID and secret, from the Cloud console under Identity and Access Management → API Access.

---

# Running it

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\CitrixUsageReport.ps1
```

That opens the dialog. `-ExecutionPolicy Bypass` is needed because Windows blocks scripts by default; it applies only to that one process and changes nothing on the machine.

_Dialog on on-premises: environment dropdown, Delivery Controller hostname, protocol, run-as choice, reporting windows, output folder, and the Include-in-the-report checklist._

_Dialog on Citrix Cloud: hostname, protocol and run-as are gone; Customer ID, Service Principal ID and Secret in their place, with a Citrix docs link._

Fields that cannot apply to your choice are **hidden**, not greyed out, and the layout closes up — choosing Citrix Cloud replaces the hostname, protocol and account fields with the three credential fields.

## Without the dialog

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command ".\CitrixUsageReport.ps1 -DeliveryController ddc01.example.com -Credential (Get-Credential) -NoGui -Days 30,60,90 -OutputPath C:\Reports -IncludeDeliveryGroups -IncludeTrend -ExportRawData"
```

Use `-Command` rather than `-File` when passing a credential — a credential object cannot cross the `-File` boundary. Write `-Days 30,60,90` with **no spaces**.

Citrix Cloud, fully scripted — read the secret **inline**, because a variable inside that quoted argument is expanded by the shell you typed it into and arrives empty:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command ".\CitrixUsageReport.ps1 -Environment CloudCommercial -CustomerId <id> -ClientId <spid> -ClientSecret (Read-Host 'Secret' -AsSecureString) -NoGui -Days 30,60,90 -OutputPath C:\Reports"
```

`-Environment` takes `OnPremises`, `CloudCommercial`, `CloudJapan` or `CloudGovernment`.

## With no Citrix environment at all

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\CitrixUsageReport.ps1 -DemoData -NoGui -OutputPath C:\Temp
```

Synthetic data, watermarked so it cannot be mistaken for a real report.

---

# Options

| Option | Effect |
| --- | --- |
| `-DeliveryController <host>` | The controller to read from. |
| `-Protocol Https\|Http` | How Monitor is published. Defaults to Https. |
| `-Environment <name>` | On-premises or which Citrix Cloud region. |
| `-CustomerId` / `-ClientId` / `-ClientSecret` | Citrix Cloud credentials. |
| `-Credential` | On-premises account. Defaults to the signed-in user. |
| `-Days 30,60,90` | Reporting windows. Maximum 1000. |
| `-OutputPath <folder>` | Where to write the report. |
| `-IncludeDeliveryGroups` | Per-delivery-group breakdown. |
| `-IncludeApplications` | Published applications and session types. Slower. |
| `-IncludeClientDevices` | Devices, addresses, Workspace app versions. Slower. |
| `-IncludeTrend` | Daily unique users and peak concurrency. |
| `-Anonymize` | Replace identities with `User-0001` style pseudonyms. Counts unchanged. |
| `-ExportRawData` | Also write the CSV and JSON exports. |
| `-NoGui` | Skip the dialog. |
| `-DemoData` | Synthetic data. Contacts nothing. |

## Combining several sites, or several months
See `RUNBOOK`.

| Option | Effect |
| --- | --- |
| `-ExportForMerge` | Also write `merge-export.json`, this site's contribution. Forces `-Anonymize` on. |
| `-SaltPath <file>` | The shared key file. Every export for one report must use the same one. |
| `-Merge <paths>` | Build a report from saved exports instead of contacting a site. Files or folders. |
| `-CrossForestKey None\|Upn` | Also match people by principal name. `None` by default. |
| `-ExcludeBridgeUpn <upn>` | Never bridge these principal names. |
| `-AllowUnanonymizedMergeExport` | Permit an export carrying raw identifiers. Deliberate only. |

---

# What comes out

A timestamped folder, e.g. `CitrixUsageReport-20260822-054409`:

| File | Produced by | What it is |
| --- | --- | --- |
| `CitrixUsageReport.html` | Every run | The report. Opens offline, has a Save as PDF button. |
| `usage-report.log` | Every run | The run log, credentials and tokens redacted. |
| `data.json`, `summary.csv`, `sessions.csv` | `-ExportRawData` | The analysis and raw records. |
| `daily-trend.csv` | `-ExportRawData` + `-IncludeTrend` | Daily figures. |
| `identity-map.csv` | `-Anonymize` | The decode key. **Keep — send everything else.** |
| `merge-export.json` | `-ExportForMerge` | This site's contribution to a combined report. No names. |
| `anonymization-salt.txt` | `-ExportForMerge` | The shared key, beside the output folder. **Keep this too.** |
| `consolidated-identity-map.csv` | A `-Merge` run | Which accounts were treated as one person, and why. No names. |

_Report header: summary table, unique users and peak concurrent tiles, and the start of a windows-compared bar chart._

---

# Reading the report

**Unique users and peak concurrency answer different questions.** 250 unique users over 90 days does not mean 250 people were logged on at once.

**The concurrency distribution is the more informative of the two.** “Highest at any single moment: 138” is the busiest instant; “19 times out of 20, at or below: 113” is what the site actually ran at nearly all the time. A peak far above the rest of the distribution was touched once.

**A truncation banner means the figures below it are a lower bound.** It appears when the window asked for is wider than the history Citrix retains — 90 days on Premium, 31 on Advanced, 7 otherwise. The real total is _higher_ than shown. Do not quote a truncated figure as a total.

Per-delivery-group, per-application, per-device and daily-trend sections appear in the report itself when their toggles are on.

---

# Combining several sites

Running per site and adding the totals up **overstates the user count** — anyone using two sites is counted twice. The last entry in the Environment dropdown builds one report from exports produced by other runs, and contacts nothing.

The full procedure, including the shared key file and how to check the result, is in `RUNBOOK`.
