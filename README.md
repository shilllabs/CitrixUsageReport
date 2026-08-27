# Citrix Usage Report

A single PowerShell script that reads usage data from a Citrix environment and produces a self-contained HTML report of unique users and concurrent sessions over the periods you ask for.

It is read-only. It issues GET requests against the Citrix Monitor Service OData API and changes nothing in the environment.

The report states what was measured. It makes no recommendations and draws no conclusions.

## What you need

- A **Delivery Controller** hostname. The Monitor Service runs there. A Citrix Director web server is a different role and will not work unless the two are installed on the same machine.
- To know whether that Delivery Controller publishes the Monitor Service over **HTTPS or HTTP**. This is a configurable Citrix setting; your Citrix administrator will know.
- An account with **read access to Citrix Monitor data**. A full administrator is not required.
- Windows PowerShell 5.1 or later. Nothing to install.

For Citrix Cloud instead of on-premises, you need the customer ID plus a service principal ID and secret, created in the Citrix Cloud console under Identity and Access Management, API Access, Service principals.

## Running it

Double-click, or open a terminal and run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\CitrixUsageReport.ps1
```

That opens a dialog where you fill in the details.

`-ExecutionPolicy Bypass` is needed because Windows blocks scripts by default. It applies only to that one process and changes nothing on the machine.

### Without a dialog

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command ".\CitrixUsageReport.ps1 -DeliveryController ddc01.example.com -Protocol Https -Credential (Get-Credential) -NoGui -Days 30,60,90 -OutputPath C:\Reports -IncludeDeliveryGroups -IncludeApplications -IncludeClientDevices -IncludeTrend -ExportRawData"
```

Two things to note. Use `-Command` rather than `-File` when passing a credential, because a credential object cannot cross the `-File` boundary. And write `-Days 30,60,90` with no spaces.

### Trying it with no Citrix environment

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\CitrixUsageReport.ps1 -DemoData -NoGui -OutputPath C:\Temp
```

Generates a report from synthetic data in about two minutes. It is watermarked so it cannot be mistaken for a real one.

## Options

| Option | Effect |
|---|---|
| `-DeliveryController <host>` | The Delivery Controller to read from. |
| `-Protocol Https\|Http` | How the Monitor Service is published. Defaults to Https. |
| `-Credential` | An account with Monitor read access. Defaults to the signed-in user. |
| `-Days 30,60,90` | Reporting periods. Maximum 1000. |
| `-OutputPath <folder>` | Where to write the report. |
| `-IncludeDeliveryGroups` | Break figures down per delivery group. |
| `-IncludeApplications` | Published applications and session types. Slower. |
| `-IncludeClientDevices` | Endpoint devices, addresses and Workspace app versions. Slower. |
| `-IncludeTrend` | Daily unique users and peak concurrency. |
| `-Anonymize` | Replace usernames with `User-0001` style identifiers. Counts are unchanged. |
| `-ExportRawData` | Also write CSV and JSON alongside the report. |
| `-NoGui` | Skip the dialog. |
| `-DemoData` | Synthetic data. Contacts no Citrix environment. |

## What you get

A timestamped folder containing:

- **`CitrixUsageReport.html`** — the report. Self-contained, opens offline, prints to PDF from a button in the page.
- **`usage-report.log`** — what happened during the run.

With `-ExportRawData`, also `data.json`, `summary.csv`, `sessions.csv`, and `daily-trend.csv` when `-IncludeTrend` is used.

With `-Anonymize`, an `identity-map.csv` is written as well. That file maps the pseudonyms back to real accounts and is the one file to keep rather than share.

## Reading the report

**Unique users** counts distinct users with at least one session in the period. A session that began before the period still counts, because it was still open during it.

**Concurrent sessions** are reported as a peak with the moment it occurred, and as a distribution in plain language — for example "19 times out of 20, at or below 109" alongside "Highest at any single moment: 144". The distinction matters: a site that touched 144 once behaves differently from one that sat near 140 all day.

**Business hours** figures are restricted to the configured working hours and weekdays, in a stated time zone. The report names the range and the zone it used.

**A truncation notice** appears when a requested period is longer than the history the environment still holds. Citrix grooms raw session data on a schedule set by the site's licence edition, so a 90-day request against a site that keeps 31 days cannot be answered in full. When that notice appears, the figures beside it are a lower bound and the real total is higher. Do not quote them as a total.

## If something goes wrong

**404, and the host serves a Director page.** The Monitor Service runs on a Delivery Controller, not on a Director web server. Point at a Delivery Controller.

**404 otherwise.** The Monitor Service may be published over HTTP rather than HTTPS on that host. Try `-Protocol Http`. A report generated over HTTP works normally and says so on its face.

**401.** The account cannot read Citrix Monitor data. Interactive runs default to the signed-in user; choose a different account in the dialog or pass `-Credential`.

**A certificate trust error.** The certificate presented by the Delivery Controller is not trusted by the machine running the report, usually because it was issued by an internal certificate authority whose root is not installed locally. Run it from a domain-joined machine inside the environment, or install the issuing root certificate.

**"could not be created or is not writable".** The output path is unreachable. The run stops before fetching anything rather than failing at the end.

## Privacy

The script reads session, user, machine, delivery group, catalog, connection and application records. It never writes to the environment and never transmits data anywhere — everything is written to the output folder on the machine that ran it.

`-Anonymize` replaces usernames, full names, UPNs, SIDs, domains, endpoint device names and addresses. It does not anonymise machine names, delivery group names, catalog names or published application names, which can be identifying or commercially sensitive in their own right.

Pseudonyms are assigned per run. `User-0007` in one report is not necessarily `User-0007` in the next, so do not compare pseudonyms across two reports.

## Note

Funnel Display and Public Sans are embedded under the SIL Open Font License. Citrix and Citrix Virtual Apps and Desktops are trademarks of Cloud Software Group, Inc. This is an independent tool and is not affiliated with or endorsed by Citrix.
