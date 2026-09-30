# Reference

* [1. Permissions](#1-permissions)
* [2. Network behaviour](#2-network-behaviour)
* [3. Fields read from Citrix](#3-fields-read-from-citrix)
* [4. What is written, and what anonymization covers](#4-what-is-written-and-what-anonymization-covers)
* [5. Cross-site identity](#5-cross-site-identity)
* [6. The merge export](#6-the-merge-export)

---

# 1. Permissions

**On-premises.** A Delegated Administration role with **read** access to the site. The built-in **Read Only Administrator** is the least-privileged built-in role that includes it. Full Administrator is not required. The account authenticates with Negotiate/NTLM as the signed-in user unless `-Credential` or the dialog's "use a different account" supplies another.

**Citrix Cloud (DaaS).** An API client (service principal) with a **read** scope, created under Identity and Access Management → API Access → Service principals. You need the customer ID, the client ID and the secret. The secret is shown once at creation; if it is lost or rotated, create a new one.

**Combining saved exports needs no access at all.** `-Merge` contacts no Citrix site, acquires no token and reads no directory — only the export files.

**Network.** Outbound HTTPS (or HTTP if that is how Monitor is published) to the Delivery Controller, or to the cloud endpoint for your region: `api.cloud.com` (Commercial), `api.citrixcloud.jp` (Japan), `api.cloud.us` (Government). The environment is always selected explicitly and never inferred — guessing wrong would send credentials to the wrong sovereign endpoint.

| Environment | Selected with | Endpoint |
| --- | --- | --- |
| On-premises | `-Environment OnPremises -DeliveryController <host>` | `<redacted URL><host>/Citrix/Monitor/OData/v4/Data` |
| Cloud Commercial | `-Environment CloudCommercial` | `https://api.cloud.com` |
| Cloud Japan | `-Environment CloudJapan` | `https://api.citrixcloud.jp` |
| Cloud Government | `-Environment CloudGovernment` | `https://api.cloud.us` |

---

# 2. Network behaviour

Exactly two kinds of call, and nothing else:
1. **HTTP GET** to the Monitor OData v4 endpoint, reading session and configuration data.
2. **One HTTP POST**, cloud only, exchanging client ID and secret for a bearer token. On-premises runs never make this call.

No PUT, PATCH or DELETE anywhere, and no call to any Citrix administrative or configuration API. The script cannot change a session, machine, delivery group or policy. `-DemoData` contacts nothing at all.

---

# 3. Fields read from Citrix

"Exported" means it reaches a file in the output folder. Sessions and the four lookup entities are fetched every run; Connections and the two Application entities only with their toggles.

**Sessions** — `SessionKey`, `UserId`, `MachineId`, `StartDate`, `EndDate`, `SessionType`, `IsAnonymous` all reach `sessions.csv`; `UserId` is pseudonymised by `-Anonymize`. `ConnectionState` and `LifecycleState` are requested for parity with Director and **discarded immediately** — neither reaches a file, the report, or any structure beyond the initial parse.

`IsAnonymous` is Citrix's own flag for unauthenticated kiosk-style launches, unrelated to this tool's `-Anonymize`.

Two one-row probes run first to establish retention: the oldest session of any kind, and the oldest **completed** session. The second sets the figure, because Citrix only grooms sessions that have ended — a desktop connected for months is never groomed and would otherwise claim history that does not exist.

**Users** — `Id`, `UserName`, `FullName`, `Upn` reach `identity-map.csv` and **only** that file, and only when `-Anonymize` is on. `Sid` is blanked and `Domain` set to `ANONYMIZED`. No Users field reaches `data.json`, `summary.csv`, `daily-trend.csv` or `sessions.csv`.

**Machines** — `Id` reaches `sessions.csv` as `MachineId`. `Name` is fetched but **never read**: delivery-group attribution goes `MachineId` → `DesktopGroupId` → `DesktopGroups.Name`, so no machine name appears anywhere. `CatalogId` likewise.

**DesktopGroups** — `Name` reaches the report and `data.json` when the delivery-group toggle is on. **Not anonymized.**

**Catalogs** — fetched every run and, in this version, read by nothing. Treat as requested from Citrix but not surfaced.

**Connections** (client-device toggle) — `ClientName` and `ClientAddress` are counted in memory and **discarded**; only the count of distinct values is reported. `ClientVersion` appears with its connection count. The fetch is limited by a `$filter` on `EstablishmentDate`, which is used in the filter only. If a site rejects that filter the script falls back to an unfiltered fetch and says so in the report.

**ApplicationInstances / Applications** (published-application toggle) — `PublishedName` reaches the report's application breakdown. **Not anonymized.** Launch fetches are likewise `$filter`ed on `StartDate`.

---

# 4. What is written, and what anonymization covers

The output files are listed in [README](https://github.com/shilllabs/CitrixUsageReport/blob/main/README.md). Three points matter for a review:

`data.json` is assembled field by field rather than serialised from the internal config, so it never contains credential material. It carries a `FetchWarnings` list naming any entity whose fetch came back incomplete — entity names and counts only.

`usage-report.log` runs every line through a redaction filter masking bearer tokens, client secrets and Basic-auth headers. It records the reporting workstation's hostname and the Windows account that ran the script, which is operator information, not Citrix site data.

If the run was over HTTP the report carries one plain navigation link to Citrix's TLS guidance — a link, not a fetched asset, so the report remains fully self-contained.

**Anonymization replaces:** usernames, full names, UPNs, SIDs (blanked), domain (`ANONYMIZED`), session-level user IDs, client device names (`Device-N`), client IP addresses (`Address-N`).

**It does not replace:** machine (VDA) names, delivery group names, machine catalog names, published application names.

> The export is **not** a fully de-identified dataset. It removes end-user identity and leaves Citrix infrastructure and application naming intact — and a delivery group named after an executive, or an application whose name reveals a line of business, is identifying or commercially sensitive in its own right.

**Pseudonyms are assigned per run**, in sorted order of whichever identifiers appear. `User-0007` in two reports is almost certainly two different people. Deriving them from the username would make them reversible by anyone holding a list of candidate usernames.

---

# 5. Cross-site identity

The problem: `Users.Id` is site-local, so the same person is a different row in every site. Three candidate identifiers were measured against live tenants:

| Candidate | Result |
| --- | --- |
| `Users.Id` | Site-local. One person was id 2 on-premises and id 16 in the same organisation's cloud tenant. |
| `Domain\UserName` | Not a definitive match without further logic from the customer. |
| `UPN` | Not a definitive match without further logic from the customer. |
| `Sid` | Stable across sites and unique within one. **Chosen as the best match for unique identity.** |

The SID is read from what Monitor already recorded. The tool never queries Active Directory and needs no directory permissions.

**Which SIDs qualify:** directory principals only — `S-1-5-21-…` (AD) and `S-1-12-1-…` (Entra ID). Anything else gets **no key at all** rather than one that looks valid. One measured edge case: a local VDA account carried a domain-*shaped* SID, because a machine's own SID is structurally identical to a domain's. It receives a key but cannot collide with a real person, since the machine portion differs from every domain's.

**Three derived keys** travel in every export — HMAC-SHA256 under the customer's salt, truncated to 128 bits, distinct labels so keys from different spaces cannot collide:

| Key | From | Label |
| --- | --- | --- |
| `userKey` | canonical SID | `citrix-usage-report/user/v1` |
| `forestKey` | the SID's domain portion | `citrix-usage-report/forest/v1` |
| `bridgeKey` | lowercased UPN | `citrix-usage-report/bridge-upn/v1` |

**Why HMAC and not a plain hash.** A domain SID prefix is not secret and RIDs allocate from a small dense range, so `SHA256(sid)` would let anyone holding an "anonymized" export enumerate the whole population by hashing candidates. The salt makes that infeasible, and the salt never leaves the customer.

**The UPN bridge** is opt-in and cannot be made safe, only careful: one person with accounts in two forests, and two different people sharing a principal name, produce byte-identical evidence. Protections: a UPN mapping to two principals within any one site has its bridge key withheld everywhere; groups spanning more than two keys or forests are flagged and still joined rather than silently dropped; `-ExcludeBridgeUpn` removes one name; the run log records whether it was used.

---

# 6. The merge export

`merge-export.json` is the only file designed to **leave** the site that produced it. It carries identity and time only:

| Field | Contents |
| --- | --- |
| `users[]` | `userKey`, `forestKey`, `bridgeKey`, `upnAmbiguousInSite` — all derived, no names |
| `sessions[]` | `sessionKey` (the site's own GUID), `userKey`, start, end, `isAnonymous` |
| `site` | environment label, cloud flag, customer id, registration hosts, delivery group and catalog ids, zones, product version, clock skew |
| windows | `windowStartUtc`, `windowEndUtc`, `historyStartUtc` |
| integrity | `schemaVersion`, `toolVersion`, `exportedUtc`, `anonymized`, `saltFingerprint`, `ambiguousUpnCount` |

> Note what `site` contains: **registration hostnames, a Citrix Cloud customer id, and for a cloud tenant the zone list**, which can name resource locations. Infrastructure naming rather than user data, but identifying, and worth knowing before the file is handed over.

`windowStartUtc` is what was **requested**; `historyStartUtc` is where the site's data actually began. Consolidation needs both — three exports each asking for 90 days look identical whether every site retained 90 days or one retained 10. Added without a schema bump because it is additive: older readers ignore it, and its absence is treated as **unconfirmed**, never as full coverage.

Coverage is judged **per site**, unioning that site's exports: six monthly exports cover six months, while a month nobody exported is a hole that coverage stops at.

`identity-map.csv` gained a `UserKey` column so the two files can be joined locally. It discloses nothing further — that file already holds the names and never leaves the machine — but without it the export and the map share no column and a combined count cannot be checked at all.

`consolidated-identity-map.csv` carries `Pseudonym`, `ConsolidatedKey`, `UserKey`, `Site`, `SourceFile`, `MatchedBy`. No names.
