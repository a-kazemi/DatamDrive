# SharePoint On-Prem Sync Client — Product & Engineering Instruction

**Working name:** DatamDrive (an open-source OneDrive replacement for on-premise SharePoint)
**Status:** Plan / instruction — not yet implemented
**Author:** CEO-review session, 2026-06-30
**Decisions locked:** True desktop sync client · Electron + Node/TypeScript · NTLM/Kerberos auth first · SharePoint 2016 / 2019 / Subscription Edition · Open-source

---

## 1. What we are building (and what we are NOT)

OneDrive's value is **not** browsing files. On-prem SharePoint already ships a web UI that lists
sites and libraries. OneDrive's value is the **sync engine**: files appear in Windows Explorer,
sync in the background, work offline, and open/save in native Word/Excel with automatic upload.

So **the sync engine is the product.** The file browser is a thin shell on top of it.

**We are building:** a Windows desktop agent that mirrors selected SharePoint document libraries
to a local folder, syncs both directions in the background, enforces SharePoint permissions
(read-only where the user has no write access), resolves conflicts safely, and lets users open
files natively and save them straight back to SharePoint.

**We are NOT building:** another web file list. The existing `portal/` app already does that.
We reuse its *concepts* (PnPjs calls, the permission model), not its web shell.

---

## 2. The hierarchy and the permission rule (your core requirement)

```
Site Collection (root web)
└── Sub-site (Web)              ── recursive, any number of levels
    └── Sub-site (Web)
        └── Document Library (List, BaseTemplate 101)
            └── Folder
                └── Folder
                    └── Document (file + SharePoint metadata)
```

**Enumeration:**
- Sub-sites: recurse `web.webs` from the site the user entered.
- Libraries: `web.lists` filtered to `BaseTemplate === 101` (document libraries).
- Folders/files: `list.rootFolder.folders` / `.files`, walked on demand.

**Permission gating (the rule you described):**
For every web, list, folder, and item, read `EffectiveBasePermissions`:

```
has ViewListItems  → show it (sync DOWN allowed)
has AddListItems / EditListItems → editable (sync UP allowed)
no write permission → mount READ-ONLY: download only, block local edits, lock icon
no read permission  → do not show, do not sync
```

A read-only node still syncs **down** so the user sees current content, but any local edit is
rejected and surfaced ("you don't have permission to change this file"). This must be enforced
in the sync engine, not just hidden in the UI — a user can edit the local file in Explorer
regardless of what the UI shows.

```
  ┌────────────────────────────────────────────┐
  │ For each node (web / library / folder / item)│
  ├────────────────────────────────────────────┤
  │ EffectiveBasePermissions?                   │
  │   contains ViewListItems? ──no──► hide, skip│
  │            │ yes                            │
  │   contains Edit/AddListItems? ─no─► READ-ONLY mount (down-only)
  │            │ yes                            │
  │            ▼                                │
  │        READ-WRITE mount (two-way)           │
  └────────────────────────────────────────────┘
```

---

## 3. Architecture

```
┌─────────────────────────── Electron App (DatamDrive) ───────────────────────────┐
│                                                                                 │
│  RENDERER (React, reuse portal/ components)        MAIN PROCESS (Node)          │
│  ┌──────────────────────────────┐                 ┌───────────────────────────┐ │
│  │ • Setup wizard (site URL,    │   IPC           │ Sync Engine (background)  │ │
│  │   pick libraries to sync)    │ ◄────────────►  │  • Auth (node-sp-auth)    │ │
│  │ • Sync status / activity log │                 │  • PnPjs v4 (@pnp/sp)     │ │
│  │ • Conflict resolution prompts│                 │  • Change poller (tokens) │ │
│  │ • Settings                   │                 │  • Local watcher (chokidar)│ │
│  └──────────────────────────────┘                 │  • Upload/Download queue  │ │
│                                                    │  • Conflict resolver      │ │
│  TRAY ICON (status, pause, quit) ◄──────────────  │  • State store (SQLite)   │ │
│                                                    └───────────┬───────────────┘ │
└──────────────────────────────────────────────────────────────┼─────────────────┘
                                                                │ HTTPS + NTLM/Kerberos
                                                                ▼
                                              On-prem SharePoint (2016/2019/SE)
                                                  _api/web, _api/lists, GetChanges
```

**Key architectural decisions:**

1. **All SharePoint calls run in the Node main process, not the renderer.** Node does NTLM
   natively, so there is **no CORS** and **no `sp-rest-proxy`**. The dev proxy in `proxy/`
   stays a dev tool only; the shipped product talks to SharePoint directly.

2. **Auth: prefer zero stored credentials.** First try the current Windows session via SSPI /
   Negotiate (Integrated Windows Auth) — the user is already domain-authenticated, so we store
   nothing. Only if SSO is unavailable do we collect domain credentials and store them in
   **Windows Credential Manager** (`keytar` / DPAPI). Never write a password to disk or app config.

3. **State store: SQLite** (`better-sqlite3`). One row per synced item mapping:
   `serverRelativeUrl ↔ localPath ↔ etag/version ↔ lastSyncedAt ↔ dirty flag ↔ permission level`.
   This table is the source of truth for "what changed where."

4. **Change detection by polling change tokens.** On-prem SharePoint has no reliable webhooks,
   so each synced library polls `GetListItemChangesSinceToken` (or `GetChanges` with a
   `ChangeQuery`) on an interval, persisting the change token between runs. Delta only, not full
   re-scan.

5. **Files-On-Demand is a later phase.** Phase 1 syncs full local copies. True placeholders
   (gray cloud / green check in Explorer) require a native cfapi addon — see §7, Phase 3.

**Reuse from the existing repo:** PnPjs setup patterns, the permission-checking logic, metadata
column handling, and the React component library from `portal/src`.

---

## 4. Feature list — what it needs to replace OneDrive

### Core (v1 — required to honestly call it a OneDrive replacement)

1. **Login + safe credential handling** — Windows SSO first; Credential Manager fallback; never plaintext.
2. **Connect by site URL** — enter a site collection URL, validate, enumerate sub-sites/libraries.
3. **Permission-aware tree** — show only what the user can read; mark no-write nodes read-only.
4. **Selective sync** — user picks which libraries/folders sync locally (don't force the whole farm).
5. **Two-way background sync** — download + upload, delta via change tokens, runs continuously.
6. **Local file watcher** — detect local creates/edits/renames/deletes (`chokidar`) and push them.
7. **Native open / save-back** — open in Word/Excel, save, auto-upload; honor check-out/check-in.
8. **Conflict detection + resolution** — local vs remote edit since last sync → keep both safely.
9. **Read-only enforcement** — block and explain local edits to files the user can't change.
10. **Delete = SharePoint Recycle Bin** — deletions are recoverable, never hard-deleted silently.
11. **Version awareness** — never clobber server version history; SharePoint etag/version is truth.
12. **Sync status UI** — tray icon, per-file status, activity log, error surfacing, pause/resume.

### OneDrive-grade expansions (Scope Expansion — opt in per item)

| # | Expansion | Why it matters | Effort (human / CC) | Risk |
|---|-----------|----------------|---------------------|------|
| E1 | **Files-On-Demand placeholders** (cfapi native addon) | The signature OneDrive feel: files visible in Explorer without using disk until opened | ~3–4 wks / ~3–4 days | **High** |
| E2 | **Explorer overlay icons** (green check / syncing) | Trust signal users expect from OneDrive | ~1–2 wks / ~2 days | Med |
| E3 | **Explorer right-click menu** (share, version history, "always keep") | Native, no app-switching | ~1 wk / ~1 day | Med |
| E4 | **Chunked large-file upload** (StartUpload/Continue/Finish) | Files > a few MB; bypasses any proxy body cap | ~3 days / ~3 hrs | Low |
| E5 | **Multi-account / multi-site** | One agent, several SharePoint sites/farms at once | ~1 wk / ~1 day | Low |
| E6 | **Offline-editable SharePoint metadata** | Edit document columns offline, sync back — ties into your DMS | ~1–2 wks / ~2 days | Med |
| E7 | **Co-authoring / live lock indicators** | Show who has a file checked out / open | ~1 wk / ~1 day | Med |
| E8 | **Bandwidth throttling + sync schedules** | IT-friendly on constrained networks | ~3 days / ~half day | Low |
| E9 | **Auto-update** (`electron-updater`) | Ship fixes without reinstall | ~2 days / ~2 hrs | Low |
| E10 | **Admin telemetry / health dashboard** | Central view of client health for IT (sellable feature) | ~2 wks / ~2 days | Med |
| E11 | **MSI + Group Policy deployment** | Enterprise rollout to many machines | ~1 wk / ~1 day | Low |
| E12 | **ADFS/SAML + FBA auth** | Federated and extranet customers (deferred from v1) | ~1–2 wks / ~2 days | Med |

> **Recommendation:** v1 = all Core items. Then E4 + E9 (cheap, high-value), then E1 (the
> make-or-break for the OneDrive feel — spike it early, §8 Risk 1). E10/E11/E12 are what turn
> this from "a tool" into "a product companies pay attention to" — schedule them once Core is solid.

---

## 5. The sync engine — pipelines and state machine

### Download pipeline (remote → local)
```
poll change token ─► get changed items ─► for each:
   ├─ new/updated remotely AND not dirty locally ─► download, update state
   ├─ new/updated remotely AND dirty locally     ─► CONFLICT (see state machine)
   ├─ deleted remotely AND not dirty locally      ─► delete local, update state
   └─ permission removed                          ─► switch node to read-only / unmount
```

### Upload pipeline (local → remote)
```
watcher event ─► debounce ─► for each changed file:
   ├─ node is read-only          ─► reject, notify user, revert/keep-local-copy
   ├─ file checked out by others  ─► queue + notify, retry later
   ├─ size > threshold            ─► chunked upload (E4)
   ├─ remote unchanged since sync ─► upload, bump version, update state
   └─ remote changed since sync   ─► CONFLICT
```

### Conflict state machine (default = keep both, never lose data)
```
        ┌──────────┐  local edit + remote edit since lastSync   ┌────────────┐
        │ IN SYNC  │ ──────────────────────────────────────────►│  CONFLICT  │
        └──────────┘                                            └─────┬──────┘
             ▲                                                        │
             │ resolved                                               ▼
             │                        ┌──────────────────────────────────────────┐
             └────────────────────────│ Keep both:                               │
                                      │  • download remote → original name        │
                                      │  • rename local → "name (conflict — user  │
                                      │    — date).ext" and upload it             │
                                      │  • log + notify; user merges manually     │
                                      └──────────────────────────────────────────┘
```

---

## 6. Auth & credential security (NTLM/Kerberos first)

```
Setup wizard
  └─ user enters site URL
       ├─ try Integrated Windows Auth (SSPI/Negotiate) using current session
       │     └─ success ─► store NOTHING, call _api/web to confirm ─► done
       └─ SSO unavailable ─► prompt domain\user + password
             └─ validate against _api/web
                   └─ store in Windows Credential Manager (keytar/DPAPI), NOT on disk
```

- Library: `node-sp-auth` (the same engine `sp-rest-proxy` already uses) supports on-prem NTLM
  and ADFS. Confirm it can consume the current Kerberos ticket; if not, use a Negotiate/SSPI
  module so we store zero credentials in the SSO path.
- **Never** persist credentials in `localStorage`, app config, or the SQLite store.
- v1 ships NTLM/Kerberos only. ADFS/SAML + FBA are E12.

---

## 7. Phased roadmap

```
Phase 0 — Foundations (spikes)
  • Electron shell + tray + IPC skeleton
  • Auth: validate NTLM/Kerberos against a real on-prem farm; confirm SSO (store-nothing) path
  • PnPjs v4 verified against SP 2016/2019/SE from the Node main process (no proxy, no CORS)
  • SQLite state store schema
  • SPIKE E1: prove a Node↔cfapi native addon can create one placeholder file  ◄ make-or-break

Phase 1 — Read path (down-only, useful on its own)
  • Permission-aware site/sub-site/library/folder tree
  • Selective sync picker
  • Full-copy download sync + change-token polling
  • Tray status + activity log

Phase 2 — Write path (two-way, the real sync client)
  • Local watcher → upload queue
  • Check-out/check-in, version-safe upload, chunked large files (E4)
  • Conflict resolver (keep both)
  • Delete → Recycle Bin
  • Read-only enforcement on edits

Phase 3 — The OneDrive feel
  • Files-On-Demand placeholders (E1), Explorer overlays (E2), context menu (E3)

Phase 4 — Product / enterprise
  • Auto-update (E9), MSI + GPO (E11), admin telemetry (E10),
    multi-account (E5), ADFS/FBA (E12), offline metadata (E6)
```

Phase 1 alone is shippable value (one-way "always have the latest files locally"). Phase 2 makes
it a true sync client. Phase 3 makes it *feel* like OneDrive.

---

## 8. Risks & open decisions (resolve early)

1. **[HIGH] Files-On-Demand on Electron.** Real placeholders need a native cfapi (C++) N-API
   addon. **Spike this in Phase 0** before committing the architecture. If it proves too costly,
   the fallback is full-copy selective sync (still a real product) and placeholders become a
   "v2 differentiator." Don't discover this late.

2. **[MED] Change detection without webhooks.** Polling change tokens per library. Tune the
   interval (responsiveness vs server load) and test with many libraries / large lists.

3. **[MED] Credential storage.** Prefer the store-nothing SSO path (current Windows Kerberos
   ticket). Stored-credential fallback must be Credential Manager/DPAPI only. Confirm
   `node-sp-auth` supports the SSO ticket; if not, pick a Negotiate module.

4. **[MED] Windows path limits & filename rules.** Local paths > 260 chars (enable long-path
   support); SharePoint-illegal characters and URL length (~400) must be sanitized/blocked with
   a clear message, never a silent failure.

5. **[LOW] Large files & the old proxy cap.** The dev proxy caps body size at 4 MB — irrelevant
   because the shipped client doesn't use it, but uploads still need the chunked API (E4) for
   big files.

6. **[DECISION] Sync conflict default.** Plan assumes "keep both." Confirm this matches user
   expectation vs "last-writer-wins + version history." Keep-both is safest and matches OneDrive.

7. **[DECISION] Open-source licensing.** "No license / open-source, maybe on Git" — pick an
   actual license (MIT for max adoption, or AGPL if you want to stop competitors from selling a
   closed fork). This is a one-way-ish door for contributors; decide before the first public push.

---

## 9. Shadow paths every flow must handle (no silent failures)

| Flow | nil / empty | upstream error | edge case |
|------|-------------|----------------|-----------|
| Enumerate site | site URL unreachable → clear error, retry | 401/403 → re-auth prompt | site has 0 sub-sites → empty-state, not crash |
| Download | item deleted mid-sync → skip + log | network drop → resume queue | disk full → pause + alert |
| Upload | file deleted locally mid-upload → cancel | checked out by other → queue + notify | rename during upload → re-resolve path |
| Auth | SSO ticket expired → silent renew → prompt | server down → retry w/ backoff | password changed in AD → re-prompt, update vault |
| Permission | revoked remotely → unmount/read-only | partial perms → per-node, not all-or-nothing | inherited vs unique perms → check effective, not assigned |

---

## 10. Definition of done (v1 / Phase 2)

- A domain user installs the app, enters a site URL, authenticates via Windows SSO (no password stored).
- They pick two libraries; one they can edit, one they can only read.
- Both appear under a local DatamDrive folder; the read-only one rejects local edits with a clear message.
- Editing a file in Word and saving uploads a new version to SharePoint; version history is intact.
- Editing the same file in two places produces a "keep both" conflict copy — no data lost.
- Deleting a file locally moves it to the SharePoint Recycle Bin.
- Killing the network mid-sync and restoring it resumes cleanly with nothing lost.
- The tray shows sync status and a readable activity/error log throughout.
```