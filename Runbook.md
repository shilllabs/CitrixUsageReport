# Runbook

* [1. Setup](#1-setup)
* [2. Running it](#2-running-it)
* [3. What gets written, and what to send back](#3-what-gets-written-and-what-to-send-back)
* [4. Troubleshooting](#4-troubleshooting)
* [5. Retention — read before trusting a 90-day number](#5-retention--read-before-trusting-a-90-day-number)
* [6. Combining several sites](#6-combining-several-sites)
* [7. Checking a combined number yourself](#7-checking-a-combined-number-yourself)
* [8. FAQ](#8-faq)

---

# 1. Setup

**Prerequisites:** PowerShell 5.1 (built into Windows 10/11); network access to a Delivery Controller or the Citrix Cloud endpoint for your region; Monitor read access ([REFERENCE](https://github.com/shilllabs/CitrixUsageReport/blob/main/Reference.md)); the `.ps1` file, saved anywhere.

1. **Execution policy** — a default client refuses to run any script, failing with a message naming `about_Execution_Policies`. That is normal and says nothing about this script.

Run it this way:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\CitrixUsageReport.ps1
```

`-ExecutionPolicy Bypass` applies to **that one process only** — no machine setting changes, nothing persists, no other script is affected. That is the form a security reviewer will want to see. (`Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` is the persistent alternative; most customers do not need it.) **Double-clicking does not work** — Windows opens `.ps1` files in an editor.

Try it safely first, contacting nothing:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\CitrixUsageReport.ps1 -DemoData -NoGui -Days 30,60,90 -OutputPath .\Demo
```

---

# 2. Running it

Run with no parameters and the dialog opens. Fields that cannot apply are hidden and the layout closes up.

| Field | Notes |
| --- | --- |
| **Environment** | On-premises, one of three Citrix Cloud regions, or **Combine saved exports** (contacts nothing — [section 6](#6-combining-several-sites)). |
| **Delivery Controller hostname** | On-premises only. A full URL can be pasted; scheme and path are stripped. |
| **Protocol** | Https (default) or Http. Typing a scheme into the hostname updates this. |
| **Customer ID / Client ID / Client secret** | Cloud only. |
| **Run as** | On-premises only. Signed-in user, or a different account. |
| **Reporting windows** | Defaults to `30,60,90`. |
| **Output folder** | Where the report is written. |
| **Include in the report** | The toggles below. |

If something is wrong, a message appears above the buttons and the dialog stays open. Otherwise the report opens in your browser when it finishes.

| Toggle | Adds | Slower? |
| --- | --- | --- |
| Delivery group breakdown | Per-group unique users and peak concurrency | No |
| Session types and published applications | Desktop vs. application split, per-application table | **Yes** |
| Client devices and Workspace app versions | Device and address counts, app versions | **Yes** |
| Daily activity trend | Day-by-day chart | Slight |
| Anonymize usernames | Pseudonyms in place of identities | No |
| Export raw data | `data.json`, `summary.csv`, `sessions.csv`, `daily-trend.csv` | No |
| Demo mode | Synthetic data, no Citrix contact | N/A |
| Export this site for combining | `merge-export.json`. Forces anonymization on | No |
| Also match people by principal name | Only when combining. **Off by default** — read [section 6](#6-combining-several-sites) | No |

**HTTP vs. HTTPS.** Whether Monitor is published over HTTP or HTTPS is a setting on the controller's IIS site; your Citrix administrator knows. HTTPS is the default for this script and the right answer wherever available. Choosing Http still completes, but the report and log both carry a visible notice that the connection was unencrypted, and that notice travels with the report.

1. By default, the CVAD Monitor service is configured to use HTTP but administrators can configure this to leverage an installed TLS certificate. More detail on this step can be found here: [https://developer-docs.citrix.com/en-us/monitor-service-odata-api/on-prem-odata.html](https://developer-docs.citrix.com/en-us/monitor-service-odata-api/on-prem-odata.html)

---

# 3. What gets written, and what to send back

Each run creates a timestamped folder. Which files appear depends on the toggles — the full list is in [README](https://github.com/shilllabs/CitrixUsageReport/blob/main/README.md).

> **Send back everything except** `identity-map.csv` **and** `anonymization-salt.txt`**.** Those two stay with you. `consolidated-identity-map.csv` holds no names and is safe to send, though it is only useful to you — see [section 7](#7-checking-a-combined-number-yourself).

`identity-map.csv` maps each `User-NNNN` pseudonym back to the account it replaced: `Pseudonym`, `UserKey`, `RealUserId`, `UserName`, `FullName`, `Upn`. `UserKey` is filled only when the run also produces a merge export. An account that appears in session data but has since been removed still gets a row, with the name columns blank.

With anonymization off there is no map to hold back, and nothing is de-identified. Anonymization never covers machine names, delivery group names or published application names — see [REFERENCE](https://github.com/shilllabs/CitrixUsageReport/blob/main/Reference.md).

---

# 4. Troubleshooting

Every failure prints a message, and the same message is in `usage-report.log`.

| Message | What it means |
| --- | --- |
| **401, "authenticated as the signed-in user"** | No account supplied and the signed-in one was rejected. Grant it Monitor read, or use **Use a different account** / `-Credential`. |
| **401, "Check the credentials supplied"** | The account you supplied was rejected. In the cloud this means the _Monitor_ call failed after a token was obtained — recheck the API client's scope. |
| **"Citrix Cloud rejected the credentials"** | The cloud sign-in itself failed: wrong customer ID, client ID or secret, or the service principal was revoked or its secret rotated. Secrets are shown once — generate a new one if lost. |
| **403** | Credentials valid, no Monitor read access. |
| **404 on the Monitor OData endpoint** | The host answers but not with the Monitor API. The message names the cause it found: a **Director web server** (point at a Delivery Controller), or **HTTPS not published** (retry with `-Protocol Http`). Neither — confirm the host really is a controller. |
| **TLS certificate trust failure** | This machine does not trust the controller's certificate, usually an internal CA whose root is missing. Run from a domain-joined machine, or install the root. There is deliberately no flag to bypass certificate validation. |
| **"Query for <Entity> failed after N attempts"** | Connection-level failure, most often DNS. Transient 408/429/5xx are retried up to 5 times first, so this means the retries were exhausted. |
| **"Could not obtain a Citrix Cloud token"** | Same family, at the first step. Check outbound HTTPS to the cloud endpoint. |
| **"No sessions were returned"** | Not an error — the run completed with zero sessions. Either access is scoped too narrowly, or the site genuinely had no activity. Check Director for the same period. |

---

# 5. Retention — read before trusting a 90-day number

Citrix keeps raw session history for **90 days** on Premium, **31** on Advanced, **7** on other editions. Only _ended_ sessions are groomed.

Before fetching, the script asks how far back data actually goes and compares it to the windows you asked for. A window wider than the retained history gets a banner beside the figures it affects:

> **This 90-day window is truncated.** This site retains only 31 day(s) of raw session history … Treat them as a **lower bound**, not a true count.

That reflects what Citrix groomed away before the script ran. The real number is _at least_ the figure shown, and cannot be recovered retroactively — only the site's edition and retention setting (`Set-MonitorConfiguration`, Premium only) change what is kept, and only from that point forward.

No banner on a window means the figures are true counts.

**The fix for short retention is to run monthly and combine** — see the next section.

To save the report as PDF, open it and click **Save as PDF**, or press Ctrl+P.

---

# 6. Combining several sites

Running per site and adding the totals up **overstates the user count**, because anyone using two sites is counted twice — as is one site reported on two dates. Measured on three real sites: a naive sum of 27 records was actually 25 people, or 23 once cross-forest accounts were matched. Nothing in the individual reports looked wrong.

No site has to be reachable from one machine, at any point.

## At each site: export its contribution

Run as usual, and also:
1. Tick **Export this site for combining**.
2. Choose the **shared key file** — at the **first** site press **Create new** and note the fingerprint; at **every other** site press **Use existing** and pick that same file, copied over. The fingerprint must match.

Each run then writes `merge-export.json` beside its report.

> **The key file is the whole arrangement.** It is what makes one person resolve to the same pseudonym everywhere, which is what lets them be counted once. Exports made under different keys can never be combined — every shared person is counted twice and nothing in the report looks wrong. Losing it permanently ends the ability to combine future exports with existing ones. It is never sent anywhere.

Ticking the box forces anonymization on: this is the file that travels between sites and must not carry raw identifiers.

## On one machine: combine them

Collect the `merge-export.json` files — the only thing that needs moving.
1. Choose **Combine saved exports from several sites (no connection)**.
2. **Add files...** for each export, or **Add folder...** to search.
3. Read the status line: it names each site, its user count and its key fingerprint, and refuses the set outright for disagreeing keys or mixed anonymization rather than producing a wrong total.
4. **Run Report**.

```powershell
powershell -ExecutionPolicy Bypass -File .\CitrixUsageReport.ps1 `
    -Merge "C:\exports" -Days 30,90,180 -OutputPath "C:\Reports"
```

## Running monthly

A site groomed at 31 days still yields a year-long picture if exports are taken before data ages out and kept. No database, no state between runs.

* **Sessions deduplicate themselves.** Session keys are GUIDs, so a session re-reported by two overlapping monthly exports counts once. One person across twelve exports is one person.
* **A repeated site is not an error.** Two exports sharing delivery group identifiers are recognised as one site.
* **Windows are measured from when the consolidation runs**, not from the data. Six months of exports with the default `30,60,90` gives the last 90 days. Ask for `-Days 30,90,180`.
* **Coverage is judged per site.** A site exported monthly for six months covers six months, even though no single export covers more than thirty days. A month nobody exported is a genuine hole: coverage stops there and every window past it is reported as a lower bound naming that site.

**Keep:** every `merge-export.json` indefinitely — they are the only record of a period once Citrix has groomed it; **the same key file throughout**; and each site's `identity-map.csv`, on the site that produced it.

## Different AD forests

By default people are matched by **security identifier**, which is exact — the same account in two sites of one forest merges with nothing switched on.

Separate forests issue different identifiers to the same person, so **Also match people across directories by user principal name** falls back to matching on name.

Leave it off unless you know your estate spans forests, and understand the limit:

> One person with accounts in two forests, and two different people sharing a principal name, produce **identical evidence**. Nothing in the data distinguishes them.

Not hypothetical — `@mil` suffixes span forests by design in government estates, and a UPN collision inside a single tenant was measured twice. What protects you: a UPN mapping to two principals _within_ any one site is withheld everywhere; groups spanning more than two keys or forests are flagged and still joined rather than silently dropped; `-ExcludeBridgeUpn` removes one recognised false match; and the run log records whether it was used.

## What a combined report cannot show, and what it warns about

The merge export carries **identity and time only**, so there is no per-delivery-group, per-application or per-device breakdown — those stay in each site's own report, and the Summary reads "Not carried in merge exports" rather than `0`.

Three notices appear only on combined reports:
* **"N of M sites do not cover this window"** — the set spans the window but not every site does, so the total is a lower bound. Different from truncation, which means one site's own history is short.
* **"Coverage unconfirmed"** — an export does not record how much history its site held. Re-export those sites to resolve it.
* **Coverage gaps** — two 90-day exports 120 days apart leave 30 days nothing covers.

---

# 7. Checking a combined number yourself

The combined report contains **no names** — that is what makes it safe to send. To satisfy yourself the number is right, join two files on your own machine:
* `identity-map.csv` — per site, has `UserKey` alongside the real `UserName`, `FullName`, `Upn`. Never leaves that site.
* `consolidated-identity-map.csv` — beside the combined report: `Pseudonym`, `ConsolidatedKey`, `UserKey`, `Site`, `SourceFile`, `MatchedBy`.

`UserKey` is the same in both, so joining them turns the anonymous total back into real accounts without a name ever entering a file you hand over.

| `MatchedBy` | Meaning |
| --- | --- |
| `SingleSite` | One site only. Nothing merged. |
| `Identifier` | Same security identifier in more than one site. Exact. |
| `UpnBridge` | Different identifiers, merged on name. **Weaker — check by hand.** |
| `Identifier+UpnBridge` | Both applied. |

**Check the** `UpnBridge` **rows.** An identifier match cannot be wrong; a name match can. Resolve them to real accounts and decide whether they are the same person. If not, drop the name match or exclude that name with `-ExcludeBridgeUpn "someone@example.com"`.

```powershell
$combined = Import-Csv .\consolidated-identity-map.csv
$people   = Get-ChildItem -Recurse -Filter identity-map.csv | ForEach-Object { Import-Csv $_.FullName }

$combined | ForEach-Object {
    $who = $people | Where-Object UserKey -eq $_.UserKey
    [pscustomobject]@{
        Pseudonym = $_.Pseudonym; Site = $_.Site; MatchedBy = $_.MatchedBy
        Account   = $who.UserName; Upn = $who.Upn
    }
} | Sort-Object Pseudonym | Format-Table -AutoSize
```

Rows sharing a `Pseudonym` are the people counted **once**.

`SourceFile` is there because the site label does not always distinguish two sites: a run on the Delivery Controller itself reports **On-premises CVAD (localhost)** whichever site it is.

---

# 8. FAQ

**Does it send anything anywhere?** No. It contacts your Citrix site and writes files locally. No telemetry, no upload. Sending anything is a manual act you choose.

**Does `User-0007` mean the same person in two reports?** **No.** Pseudonyms are assigned fresh each run, in sorted order of whichever identifiers appear, so anyone added or removed between runs shifts every pseudonym after them. That is deliberate: deriving them from the username would make them reversible by anyone holding a list of candidate usernames. Each report is internally consistent with its own map. To follow one person across two reports, resolve each pseudonym through its own map first, then compare the real accounts.

**Can the pseudonyms be reversed by whoever receives the report?** Not from the report. The cross-site keys are HMAC-SHA256 under your salt, which never leaves your estate.

**Is the concurrency figure exact?** Yes — a sweep-line over session intervals, hand-verified against raw OData three times. A session ending exactly as another begins is not counted as concurrent.

**Does a failed or instant session count as a user?** Currently **no**, and it is an open question. A session whose start equals its end contributes no time and is discarded — correct for concurrency, arguably wrong for licensing since the person did log on. One measured 90-day window reported 13 unique users where 14 had connected.

**Why is the report's count lower than the rows in `identity-map.csv`?** They answer different questions. The map lists every account the site knows; the report counts those with sessions in the window. 18 rows against a reported 9 is not a discrepancy.

**Does the report list the individual users?** No — counts only, per-site and combined. To see who was counted, tick **Export raw data** and join `sessions.csv` (`UserId`) to `identity-map.csv` (`RealUserId`), filtered to the window.

**What if I use two different key files by accident?** The merge refuses the whole set and says why, rather than combining them and counting every shared person twice.

**How do I know two sites used the same key?** Each export carries a fingerprint, shown in the dialog for every file you add. Same fingerprint, same key.

**Can I schedule it?** Yes — run non-interactively with `-ExportForMerge` and a fixed `-SaltPath` and you accumulate a combinable series. That is the recommended answer to short retention.

**It printed "Canceled." and nothing else.** On 1.5.0 and 1.6.0 that was a defect — every Run reported canceled. Fixed in 1.6.1. On later builds it means the dialog was closed. Check the version in the first lines of the run log.

**Can this prove my customer gave us data from every site?** **No, and it does not claim to.** The tool reports on the sites it is given and cannot know how many exist, so three sites out of four produces a report that looks complete and is not. Where license server records exist the set of sites can be cross-checked against them; in air-gapped estates there is no such check. What it does establish is what the presented sites show, with an auditable method: an exact concurrency calculation, explicit retention caveats, named refusals rather than silent merges, and a validation path the customer can run themselves.
